# Import the necessary module
import pandas as pd
import numpy as np
import os
from PyQt5 import QtWidgets
from PyQt5 import QtGui
import chardet
import datetime
import shutil

class myApp(QtWidgets.QWizard):
    def __init__(self, parent=None):
        super(myApp, self).__init__(parent)

        """ Initialise Environment (If not yet initialised) """
        self.envInit()

        self.mainPage = MainPage()
        self.addPage(self.mainPage)


        self.setPixmap(QtWidgets.QWizard.BannerPixmap, QtGui.QPixmap(os.getcwd()+"/CASSCare.png"))

        self.setWindowTitle("Extra Mileage Loading Calculator")
        self.setWindowIcon(QtGui.QIcon(r'icon.ico'))
        self.setWizardStyle(1)

        self.setButtonText(self.FinishButton, "Run!")
        #self.button(QtWidgets.QWizard.FinishButton).clicked.connect(self.mainPage.saveFile)
        self.button(QtWidgets.QWizard.FinishButton).clicked.connect(self.runProgram)

        self.setOption(self.NoBackButtonOnStartPage, True)

        self.errorCode = 0
        self.mismatchInfo = ""
        self.missingIndex = ""
        self.attacheFilePath = ""
        self.mileageFilePath = ""
        self.outputFilePath = ""
        self.arcDirectory = ""

    """Predefined functions"""
    """envInit - initialise working directories"""
    def envInit(self):
        if not(os.path.isdir("DATA/current/input_files/")):
            os.makedirs("DATA/current/input_files",exist_ok=True)
        if not(os.path.isdir("DATA/current/output_file/")):
            os.makedirs("DATA/current/output_file",exist_ok=True)
        if not(os.path.isdir("DATA/archive/")):
            os.makedirs("DATA/archive",exist_ok=True)


    """errorMsg - Pop Up Window Box to show the error Message"""
    def errorMsg(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        if self.errorCode == 1:
            msg.setText("You have choose to run in Custom Mode, but you haven't choose any file or folders.")
        elif self.errorCode == 2:
            msg.setText("Invalid file name or file path has been selected.")
        elif self.errorCode == 3:
            msg.setText("The Selected Attache Data File is Invalid.")
        elif self.errorCode == 4:
            msg.setText("The selected column is not pure text.")
        elif self.errorCode == 5:
            msg.setText("You have mismatched mileage record, the total mileage calculated from per visit is not the same as the one recorded in Attache import data file.")
            msg.setDetailedText(self.mismatchInfo)
        elif self.errorCode == 6:
            msg.setText("The input file(s) not exist. Please check the DATA\\current\\input_files folder")
        elif self.errorCode == 7:
            msg.setText("The Selected Mileage Data File is Invalid.")
        elif self.errorCode == 8:
            msg.setText("There are missing entries in Attache File.")
            msg.setDetailedText(self.missingIndex)
        msg.setWindowTitle("Oppos...")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()

    """successMsg - Pop Up Window Box to show the File Saved Success Message"""
    def successMsg(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setText("Operation Complete. File Saved at: "+os.path.normpath(os.path.abspath(self.outputFilePath)))
        msg.setWindowTitle("Success!")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()

    """sucArcMsg - Pop Up Window Box to show the File Achived Success Message"""
    def sucArcMsg(self):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setText("File Archived. File Saved at: "+os.path.normpath(os.path.abspath(self.arcDirectory)))
        msg.setWindowTitle("Success!")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()

    """capColumn - Capitalise the String data in a specific column"""
    def capColumn(self, df, columnIdx):
        for idx in df.index:
            if not isinstance(df.iloc[idx][columnIdx],str):
                self.errorCode = 4
                self.errorMsg()
                sys.exit()
            content = df.iloc[idx][columnIdx]
            contentUpper = content.upper()
            df.set_value(idx,columnIdx,contentUpper)
        return df

    """emCodeExt - extract employee code from Attache import file"""
    def emCodeExt(self, df, colIdx):
        emCodes = []
        for idx in df.index:
            code = df.iloc[idx][colIdx]
            if not(code in emCodes):
                emCodes.append(code)
        return emCodes

    """emPayExt - extract employee salary from Attache import file"""
    def emPayExt(self, df, emCodes):
        emPays = {}
        pays = [0]*2
        for code in emCodes:
            tableSlice = df[(df[2] == code) & ((df[5] != "BACKPAY") & (df[5] != "MILEINT") & (df[5] != "MILEOUT") & (df[5] != "TRANSPORT") & (df[9] != 0.78))]
            highPay = max(tableSlice[9])
            lowPay = min(tableSlice[9])
            highPaySection = tableSlice[tableSlice[9]==highPay]
            lowPaySection = tableSlice[tableSlice[9]==lowPay]
            if (highPaySection.iloc[0][5] == "ES"):
                pays[0] = highPaySection.iloc[0][9] * 8
            elif (highPaySection.iloc[0][5] == "NS"):
                pays[0] = highPaySection.iloc[0][9] / 0.15
            elif (highPaySection.iloc[0][5] == "SAT"):
                pays[0] = highPaySection.iloc[0][9] * 2
            elif (highPaySection.iloc[0][5] == "PHLOAD"):
                pays[0] = highPaySection.iloc[0][9] / 1.5
            else:
                pays[0] = highPaySection.iloc[0][9]
            if (lowPaySection.iloc[0][5] == "ES"):
                pays[1] = lowPaySection.iloc[0][9] * 8
            elif (lowPaySection.iloc[0][5] == "NS"):
                pays[1] = lowPaySection.iloc[0][9] / 0.15
            elif (lowPaySection.iloc[0][5] == "SAT"):
                pays[1] = lowPaySection.iloc[0][9] * 2
            elif (lowPaySection.iloc[0][5] == "PHLOAD"):
                pays[1] = lowPaySection.iloc[0][9] / 1.5
            else:
                pays[1] = lowPaySection.iloc[0][9]
            emPays[code] = pays[:]
        return emPays

    """mileInCheck - Check the recorded milage data"""
    def mileInCheck(self, dfA, dfM, emCodes):
        hasMismatch = False
        self.mismatchInfo = "You have mismatched mileage record, the total mileage calculated from per visit is not the same as the one recorded in Attache import data file\n"
        for code in emCodes:
            dfA_slice = dfA[(dfA[2] == code) & (dfA[5]=="MILEINT")]
            ccA = dfA_slice[12].tolist()
            mileAcc = dfA_slice[6].tolist()
            mileAcc = [ '%.3f' % elem for elem in mileAcc ]
            mileAcc = map(float, mileAcc)
            mileageDict = dict(zip(ccA, mileAcc))
            dfM_slice = dfM[dfM["Employee Code"] == code]
            for cc in ccA:
                dfM_slice_fine = dfM_slice[dfM_slice["Client Cost Center"] == cc]
                totalMile = dfM_slice_fine["KMs"].sum()
                totalMile = round(totalMile,3)
                if~(totalMile == mileageDict[cc]):
                    diff = abs(totalMile - mileageDict[cc])
                    if diff > 0.005:
                        self.mismatchInfo = self.mismatchInfo + "Employee Code: {}, Cost Centre: {}. Attache Value: {}, Calculated Value: {}.\n".format(code, cc, mileageDict[cc], totalMile)
                        hasMismatch = True
        if hasMismatch:
            self.errorCode = 5
            self.errorMsg()
            sys.exit()


    """mileRM - Remove unrelated records"""
    def mileRM(self, dfM):
        visitDate = dfM["Visit Date"]
        theDate = visitDate.apply(lambda x: x.split(' ')[0])
        x = theDate[theDate.index[0]]

        # Added Additional Conditions to check the employee Code #
        theEmployee = dfM["Employee Code"]
        y = theEmployee[theEmployee.index[0]]
        # End of Additional Conditions #

        resultIndex = []
        for i in range(1,len(theDate.index)):
            l = [ord(a) ^ ord(b) for a,b in zip(x,theDate[theDate.index[i]])]
            m = [ord(c) ^ ord(d) for c,d in zip(y,theEmployee[theEmployee.index[i]])]
            x = theDate[theDate.index[i]]
            y = theEmployee[theEmployee.index[i]]
            if (sum(l) == 0) and (sum(m) == 0):
                resultIndex.append(theDate.index[i-1])
        resultIndex = list(set(resultIndex))
        dfM_FF = dfM.loc[resultIndex]
        return dfM_FF

    """mileBonus - give additional 5% of mileage amount to its actual travel"""
    def mileBonus(self, dfM):
        dfM_KMs = dfM["KMs"]
        dfM_KMs = dfM_KMs.apply(lambda x:x*1.30)
        dfM["KMs"] = dfM_KMs
        return dfM

    """mileInLoading - Calculate additional mileage loading"""
    def mileInLoading(self, dfA, dfM, emCodes):
        extraMileage = {}
        for code in emCodes:
            workerExtraMiles = {}
            dfA_slice = dfA[(dfA[2] == code) & (dfA[5]=="MILEINT")]
            ccA = dfA_slice[12].tolist()
            dfM_slice = dfM[dfM["Employee Code"] == code]
            for cc in ccA:
                extraMiles = []
                dfM_slice_fine = dfM_slice[dfM_slice["Client Cost Center"] == cc]
                visitTimes = dfM_slice_fine["Visit Date"]
                visitTimes = visitTimes.apply(lambda x: int(x[len(x)-5:len(x)-3]))
                ordShifts = dfM_slice_fine.loc[visitTimes>=6]
                ordShifts = ordShifts.loc[visitTimes<20]
                #ordShifts1 = dfM_slice_fine[visitTimes>=6]
                #ordShifts1 = ordShifts1[visitTimes<20]
                esShifts = dfM_slice_fine[visitTimes>=20]
                nsShifts = dfM_slice_fine[visitTimes<6]
                """ Extra Mileage for ORD work shifts """
                if ordShifts.empty:
                    ordextraMiles = []
                else:
                    ordmileages = ordShifts["KMs"]
                    ordmileages= round(ordmileages,2)
                    ordmileCandidates = ordmileages[ordmileages > 10.0]
                    ordextraMiles = (ordmileCandidates-10.0)/30.0
                    ordextraMiles = ordextraMiles.tolist()
                """ Extra Mileage for ES work shifts """
                if esShifts.empty:
                    esextraMiles = []
                else:
                    esmileages = esShifts["KMs"]
                    esmileages= round(esmileages,2)
                    esmileCandidates = esmileages[esmileages > 10.0]
                    esextraMiles = (esmileCandidates-10.0)/30.0
                    esextraMiles = esextraMiles.tolist()
                """ Extra Mileage for ES work shifts """
                if nsShifts.empty:
                    nsextraMiles = []
                else:
                    nsmileages = nsShifts["KMs"]
                    nsmileages= round(nsmileages,2)
                    nsmileCandidates = nsmileages[nsmileages > 10.0]
                    nsextraMiles = (nsmileCandidates-10.0)/30.0
                    nsextraMiles = nsextraMiles.tolist()
                extraMiles=[ordextraMiles,esextraMiles,nsextraMiles]
                workerExtraMiles[cc] = extraMiles
            extraMileage[code] = workerExtraMiles
        return extraMileage


    """addML - Add the additional mileage loading to existing Attache Data Records"""
    def addML(self, dfA, emCodes, emPays, extraMiles):
        newDF = pd.DataFrame(columns=range(0,23)) # New Data Frame for store the update records
        for code in emCodes:
            newDF = newDF.append(dfA[dfA[2] == code], ignore_index=True)
            for cc in extraMiles[code]:
                for shiftIdx in range(3):
                    shift = extraMiles[code][cc][shiftIdx]
                    listLen = len(shift)
                    if listLen > 0:
                        for miles in shift:
                            newRow = pd.Series(index=range(23))
                            newRow[0] = "T6"
                            newRow[2] = code
                            newRow[4] = "N"
                            newRow[5] = "ORD"
                            newRow[6] = round(miles,2)
                            if (shiftIdx == 0):
                                newRow[9] = emPays[code][0]
                            else:
                                newRow[9] = emPays[code][1]
                            newRow[12] = cc
                            newDF = newDF.append(newRow, ignore_index=True)
        return newDF

    """reGenData - Regenerate Attache Data Records"""
    def reGenData(self, dfA, dfM, emCodes, emPays, extraMiles):
        newDF = pd.DataFrame(columns=range(0,23)) # New Data Frame for store the update records
        for code in emCodes:
            dfM_slice = dfM[dfM["Employee Code"] == code]
            for cc in extraMiles[code]:
                dfM_slice_fine = dfM_slice[dfM_slice["Client Cost Center"] == cc]
                if round(sum(dfM_slice_fine["KMs"]),2) != 0.000:
                    newRow = pd.Series(index=range(23))
                    newRow[0] = "T6"
                    newRow[2] = code
                    newRow[4] = "A"
                    newRow[5] = "MILEINT"
                    newRow[6] = round(sum(dfM_slice_fine["KMs"]),2)
                    newRow[9] = 0.78
                    newRow[12] = cc
                    newDF = newDF.append(newRow, ignore_index=True)
            newDF = newDF.append(dfA[dfA[2] == code], ignore_index=True)
            for cc in extraMiles[code]:
                for shiftIdx in range(3):
                    shift = extraMiles[code][cc][shiftIdx]
                    listLen = len(shift)
                    if listLen > 0:
                        for miles in shift:
                            newRow = pd.Series(index=range(23))
                            newRow[0] = "T6"
                            newRow[2] = code
                            newRow[4] = "N"
                            newRow[5] = "ORD"
                            newRow[6] = round(miles,2)
                            if (shiftIdx == 0):
                                newRow[9] = emPays[code][0]
                            else:
                                newRow[9] = emPays[code][1]
                            newRow[12] = cc
                            newDF = newDF.append(newRow, ignore_index=True)
        return newDF

    """rpPayType - Replace OUTING and TRANSPORT paytype with MILEOUT and MILEIN"""
    def rpMileout(self, df):
        dfReplaced = df.apply(lambda x: x.replace("OUTING","MILEOUT"))
        dfReplaced = dfReplaced.apply(lambda x: x.replace("MILEAGE(OUTING)","MILEOUT"))
        dfReplaced = dfReplaced.apply(lambda x: x.replace("TRANSPORT","MILEOUT"))
        dfReplaced = dfReplaced.apply(lambda x: x.replace("TRANSPORT SERVICES","MILEOUT"))
        return dfReplaced

    """ arcFile - Archive the source input files and generated Attache Output Data files to the default archieve folder """
    def archiveFile(self):
        cDate = datetime.datetime.now()
        cDateString = str(cDate.year)+str(cDate.month)+str(cDate.day)
        self.arcDirectory = "DATA/archive/"+cDateString
        if not(os.path.isdir(self.arcDirectory)):
            os.mkdir(self.arcDirectory)
        shutil.copyfile(self.attacheFilePath, self.arcDirectory+"/attacheExport_"+cDateString+".csv")
        shutil.copyfile(self.mileageFilePath, self.arcDirectory+"/mileageExport_"+cDateString+".csv")
        shutil.copyfile(self.outputFilePath, self.arcDirectory+"/PAYTSHT_"+cDateString+".INP")


    """Main part of the program"""
    def runProgram(self):
        """ Extract Feild Entry """
        AttacheFileName = self.mainPage.field("AttacheFileName")
        MileageFileName = self.mainPage.field("MileageFileName")
        OutputFileName = self.mainPage.field("OutputFileName")

        """ Check User inputs"""
        if (self.mainPage.runMode == 2):
            if (not AttacheFileName) and (not MileageFileName) and (not OutputFileName): # Three Fields are all empty
                self.errorCode = 1
                self.errorMsg()
                sys.exit()
            elif (self.mainPage.inputFile1NameLineEdit == ".") or (MileageFileName == ".") or (OutputFileName == "."): # One of the Field has a None string due to Cancel btn clicked
                self.errorCode = 2
                self.errorMsg()
                sys.exit()
            elif AttacheFileName and (not MileageFileName) and (not OutputFileName): # Attache Custom File is selected, rest are empty
                self.attacheFilePath = self.mainPage.inf1PH_C[0]
                self.mileageFilePath = self.mainPage.inf2PH
                self.outputFilePath = self.mainPage.outfPH
            elif AttacheFileName and (not MileageFileName) and (OutputFileName): # Attache Custom File is selected, and Custom Output Directory is defined, rest are empty
                self.attacheFilePath = self.mainPage.inf1PH_C[0]
                self.mileageFilePath = self.mainPage.inf2PH
                self.outputFilePath = self.mainPage.outfPH_C + "/PAYTSHT.INP"
            elif AttacheFileName and (MileageFileName) and (not OutputFileName): # Attache Custom File and Mileage Custom File are selected, rest are empty
                self.attacheFilePath = self.mainPage.inf1PH_C[0]
                self.mileageFilePath = self.mainPage.inf2PH_C[0]
                self.outputFilePath = self.mainPage.outfPH
            elif (not AttacheFileName) and (MileageFileName) and (not OutputFileName): # Mileage Custom File is selected, rest are empty
                self.attacheFilePath = self.mainPage.inf1PH
                self.mileageFilePath = self.mainPage.inf2PH_C[0]
                self.outputFilePath = self.mainPage.outfPH
            elif (not AttacheFileName) and (MileageFileName) and (OutputFileName): # Mileage Custom File is selected, and Custom Output Directory is defined, rest are empty
                self.attacheFilePath = self.mainPage.inf1PH
                self.mileageFilePath = self.mainPage.inf2PH_C[0]
                self.outputFilePath = self.mainPage.outfPH_C + "/PAYTSHT.INP"
            elif (not AttacheFileName) and (not MileageFileName) and (OutputFileName): # Custom Output Directory is defined, rest are empty
                self.attacheFilePath = self.mainPage.inf1PH
                self.mileageFilePath = self.mainPage.inf2PH
                self.outputFilePath = self.mainPage.outfPH_C + "/PAYTSHT.INP"
            else:
                self.attacheFilePath = self.mainPage.inf1PH_C[0]
                self.mileageFilePath = self.mainPage.inf2PH_C[0]
                self.outputFilePath = self.mainPage.outfPH_C + "/PAYTSHT.INP"
        else:
            self.attacheFilePath = self.mainPage.inf1PH
            self.mileageFilePath = self.mainPage.inf2PH
            self.outputFilePath = self.mainPage.outfPH

        """ Check if the existance of the input files """
        if not(os.path.isfile(self.attacheFilePath)) or not(os.path.isfile(self.mileageFilePath)):
            self.errorCode = 6
            self.errorMsg()
            sys.exit()

        """ Detect the file encoding scheme """
        f = open(self.attacheFilePath,"rb")
        file = f.read()
        detectResult = chardet.detect(file)

        """ Load the Attache Payroll file into a data frame """
        df_main = pd.read_csv(self.attacheFilePath,header=None,converters={2 : str, 4 : str, 5 : str, 12 : str})

        """ Check to see if a valid Attache Data File is provided """
        self.missingIndex = "You have missing entry. Please check: \n"
        misIdxs = []
        if detectResult["encoding"] == "ascii":
            firstEntry = df_main.iloc[0][0]
        else:
            firstEntry = df_main.iloc[1][0]
        if not(firstEntry == "T6"):
                self.errorCode = 3
                self.errorMsg()
                sys.exit()
        elif not(len(list(df_main.columns.values)) == 24):
            self.errorCode = 3
            self.errorMsg()
            sys.exit()
        #misIdxs = (np.isnan(df_main.iloc[:,2]).index.tolist()) + (np.isnan(df_main.iloc[:,4]).index.tolist()) + (np.isnan(df_main.iloc[:,5]).index.tolist()) + (np.isnan(df_main.iloc[:,6]).index.tolist()) + (np.isnan(df_main.iloc[:,9]).index.tolist())+ (np.isnan(df_main.iloc[:,12]).index.tolist())
        colValues = df_main.iloc[:,2]
        misIdxs = misIdxs + colValues[(colValues == "")].index.tolist()
        colValues = df_main.iloc[:,4]
        misIdxs = misIdxs + colValues[(colValues == "")].index.tolist()
        colValues = df_main.iloc[:,5]
        misIdxs = misIdxs + colValues[(colValues == "")].index.tolist()
        colValues = df_main.iloc[:,6]
        misIdxs = misIdxs + colValues[np.isnan(colValues)].index.tolist()
        colValues = df_main.iloc[:,9]
        misIdxs = misIdxs + colValues[np.isnan(colValues)].index.tolist()
        colValues = df_main.iloc[:,12]
        misIdxs = misIdxs + colValues[(colValues == "")].index.tolist()
        misIdxs = list(set(misIdxs))
        if len(misIdxs) > 0:
            for misIdx in misIdxs:
                self.missingIndex = self.missingIndex + "Line: {}\n".format(misIdx+1)
            self.errorCode = 8
            self.errorMsg()
            sys.exit()

        """ Capitalise the Pay Type Column """
        df_main = self.capColumn(df_main,5)


        """ Remove all existing MILEINT entries """
        df_main_filtered = df_main[~(df_main[5] == "MILEINT")]


        """ Generate employee code lists """
        emCodes = self.emCodeExt(df_main,2)

        """ Extract Employee Salary """
        emPays = self.emPayExt(df_main, emCodes)

        """ Load the mileage data into a data frame """
        df_mile_org = pd.read_csv(self.mileageFilePath, converters={"Employee Code" : str}, sep="\t")

        """ Check to see if a valid Mileage Data File is provided """
        df_mile_header = list(df_mile_org.columns.values)
        if not(len(df_mile_header) == 6):
            self.errorCode = 7
            self.errorMsg()
            sys.exit()
        if (not(df_mile_header[0] == "Worker Name")) or (not(df_mile_header[1] == "Employee Code")) or (not(df_mile_header[2] == "Client Name")) or (not(df_mile_header[3] == "Client Cost Center")) or (not(df_mile_header[4] == "Visit Date")) or (not(df_mile_header[5] == "KMs")):
            self.errorCode = 7
            self.errorMsg()
            sys.exit()

        # Filter the table to remove non recorded rows
        df_mile = df_mile_org[~(df_mile_org["KMs"].isnull())]

        """ Verify if there is any error with the recorded mileage """
        checkResult = self.mileInCheck(df_main, df_mile, emCodes)

        """ Remove unrelated Mileage Records """
        df_mile_filtered = self.mileRM(df_mile)

        """ Calculate the bonus 5% mileage """
        df_mile_filtered = self.mileBonus(df_mile_filtered)

        """ Determine the additional amount of Milein needed for each worker """
        extraMiles = self.mileInLoading(df_main, df_mile_filtered, emCodes)

        """ Generate the new dataframe with additional mileage loading """
        #newDF = self.addML(df_main_filtered,emCodes,emPays,extraMiles)
        #print("run to here")

        """ Generate the new dataframe with additional mileage loading """
        newDF = self.reGenData(df_main_filtered, df_mile_filtered, emCodes, emPays, extraMiles)

        """ Replace the Pay Type """
        newDF = self.rpMileout(newDF)

        """ Re-encode the string columns using ASCII method """
        newDF[0] = newDF[0].map(lambda x: x.encode(encoding='ascii',errors='ignore').decode(encoding='ascii',errors='ignore'))
        newDF[2] = newDF[2].map(lambda x: x.encode(encoding='ascii',errors='ignore').decode(encoding='ascii',errors='ignore'))
        newDF[4] = newDF[4].map(lambda x: x.encode(encoding='ascii',errors='ignore').decode(encoding='ascii',errors='ignore'))
        newDF[5] = newDF[5].map(lambda x: x.encode(encoding='ascii',errors='ignore').decode(encoding='ascii',errors='ignore'))
        newDF[12] = newDF[12].map(lambda x: x.encode(encoding='ascii',errors='ignore').decode(encoding='ascii',errors='ignore'))

        """ Save the updated records to a csv file """
        newDF.to_csv(self.outputFilePath,header=None,index=None,encoding="ascii",float_format='%.3f')
        self.successMsg()

        """ Archive the files if requested by user """
        if self.mainPage.archive:
            self.archiveFile()
            self.sucArcMsg()



class MainPage(QtWidgets.QWizardPage):

    def __init__(self, parent=None):
        super(MainPage, self).__init__(parent)

        self.inf1PH = "DATA/current/input_files/attacheExport.csv"
        self.inf2PH = "DATA/current/input_files/mileageExport.csv"
        self.inf3PH = "DATA/current/input_files/DCW_Rates.xlsx"
        self.outfPH = "DATA/current/output_file/PAYTSHT.INP"

        self.runMode = 0
        self.archive = True

        self.setTitle("Welcome!")
        self.setPixmap(QtWidgets.QWizard.WatermarkPixmap, QtGui.QPixmap("CASSCare.png"))


        label = QtWidgets.QLabel("Here are the features of this application: \n" +
                                 "- Calculates any additional mileage load that is above 10km \n" +
                                 "- Updates the weekend rates for casual workers \n" +
                                 "- Updates the internet allowance from 1.5 to 1.25 \n" +
                                 "- Change SL to PCL \n" +
                                 "- Add \"leave loading\" to AL \n" +
                                 "- Add \"internet\" allowance to every staff\n\n" +
                                 "To start, you can choose to run the program using default settings.\n" +
                                 "Otherwise, you can choose to customise the setting.")

        label.setWordWrap(True)

        inputFile1NameLabel = QtWidgets.QLabel("Choose Your Attache Data File:")
        self.inputFile1NameLineEdit = QtWidgets.QLineEdit()
        inputFile1NameLabel.setBuddy(self.inputFile1NameLineEdit)

        inputFile1SelectionBtn = QtWidgets.QPushButton("...", parent = None)
        self.inf1PH_C = ""
        #self.inf1PH_C = QtWidgets.QFileDialog.getOpenFileName(self, None, "Choose Attache Data File..", "/home", "Comma Separated Values File (*.csv);;Attache Data File (*.INP);;All Files (*)")

        inputFile2NameLabel = QtWidgets.QLabel("Choose Your Mileage Data File:")
        self.inputFile2NameLineEdit = QtWidgets.QLineEdit()
        inputFile2NameLabel.setBuddy(self.inputFile2NameLineEdit)

        inputFile2SelectionBtn = QtWidgets.QPushButton("...", parent = None)
        self.inf2PH_C = ""

        inputFile3NameLabel = QtWidgets.QLabel("Choose Your DCW Rates file (.xlsx):")
        self.inputFile3NameLineEdit = QtWidgets.QLineEdit()
        inputFile3NameLabel.setBuddy(self.inputFile3NameLineEdit)

        inputFile3SelectionBtn = QtWidgets.QPushButton("...", parent = None)
        self.inf3PH_C = ""

        outputFileNameLabel = QtWidgets.QLabel("Select the folder for the Output File:")
        self.outputFileNameLineEdit = QtWidgets.QLineEdit()
        outputFileNameLabel.setBuddy(self.outputFileNameLineEdit)

        outputFileSelectionBtn = QtWidgets.QPushButton("...", parent = None)
        self.outfPH_C = ""

        archiveCheckBox = QtWidgets.QCheckBox("Archive Files?")
        archiveCheckBox.setChecked(True)

        self.registerField("AttacheFileName",self.inputFile1NameLineEdit)
        self.registerField("MileageFileName",self.inputFile2NameLineEdit)
        self.registerField("DCWRatesFileName",self.inputFile3NameLineEdit)
        self.registerField("OutputFileName",self.outputFileNameLineEdit)
        self.registerField("fileArchiveFlag",archiveCheckBox)

        runModeBox = QtWidgets.QGroupBox("Running Mode")

        defaultModeRadioButton = QtWidgets.QRadioButton("Default Mode")
        customModeRadioButton = QtWidgets.QRadioButton("Custom Mode")

        defaultModeRadioButton.setChecked(True)

        if defaultModeRadioButton.isChecked():
            self.runMode = 1
            self.inputFile1NameLineEdit.setEnabled(False)
            self.inputFile2NameLineEdit.setEnabled(False)
            self.inputFile3NameLineEdit.setEnabled(False)
            self.outputFileNameLineEdit.setEnabled(False)
            archiveCheckBox.setEnabled(False)
            inputFile1SelectionBtn.setEnabled(False)
            inputFile2SelectionBtn.setEnabled(False)
            inputFile3SelectionBtn.setEnabled(False)
            outputFileSelectionBtn.setEnabled(False)
            self.inputFile1NameLineEdit.setText(os.path.normpath(self.inf1PH))
            self.inputFile2NameLineEdit.setText(os.path.normpath(self.inf2PH))
            self.inputFile3NameLineEdit.setText(os.path.normpath(self.inf3PH))
            self.outputFileNameLineEdit.setText(os.path.normpath(self.outfPH))



        #defaultModeRadioButton.toggled.connect(ToggleFlag)

        #defaultModeRadioButton.toggled.connect(self.setButtonText(QtWidgets.QWizard.WizardButton(1), "run!"))

        #defaultModeRadioButton.toggled.connect(inputFile1NameLineEdit.setDisabled)
        #defaultModeRadioButton.toggled.connect(inputFile2NameLineEdit.setDisabled)
        #defaultModeRadioButton.toggled.connect(outputFileNameLineEdit.setDisabled)
        defaultModeRadioButton.toggled.connect(archiveCheckBox.setDisabled)
        defaultModeRadioButton.toggled.connect(archiveCheckBox.setChecked)
        defaultModeRadioButton.toggled.connect(inputFile1SelectionBtn.setDisabled)
        defaultModeRadioButton.toggled.connect(inputFile2SelectionBtn.setDisabled)
        defaultModeRadioButton.toggled.connect(inputFile3SelectionBtn.setDisabled)
        defaultModeRadioButton.toggled.connect(outputFileSelectionBtn.setDisabled)
        defaultModeRadioButton.toggled.connect(self.alterMode)

        archiveCheckBox.stateChanged.connect(self.alterArchive)

        inputFile1SelectionBtn.clicked.connect(self.selectAttacheFile)
        inputFile2SelectionBtn.clicked.connect(self.selectMileageFile)
        inputFile3SelectionBtn.clicked.connect(self.selectDCWRatesFile)
        outputFileSelectionBtn.clicked.connect(self.selectOutputDirc)

        self.registerField("defaultMode", defaultModeRadioButton)
        self.registerField("customMode", customModeRadioButton)

        runModeBoxLayout = QtWidgets.QVBoxLayout()
        runModeBoxLayout.addWidget(defaultModeRadioButton)
        runModeBoxLayout.addWidget(customModeRadioButton)
        runModeBox.setLayout(runModeBoxLayout)

        layout = QtWidgets.QGridLayout()
        layout.addWidget(label,0,0)
        layout.addWidget(runModeBox,1,0)

        layout.addWidget(inputFile1NameLabel,2,0)
        layout.addWidget(self.inputFile1NameLineEdit,3,0,1,2)
        layout.addWidget(inputFile1SelectionBtn,3,3,1,1)

        layout.addWidget(inputFile2NameLabel,4,0)
        layout.addWidget(self.inputFile2NameLineEdit,5,0,1,2)
        layout.addWidget(inputFile2SelectionBtn,5,3,1,1)

        layout.addWidget(inputFile3NameLabel,6,0)
        layout.addWidget(self.inputFile3NameLineEdit,7,0,1,2)
        layout.addWidget(inputFile3SelectionBtn,7,3,1,1)

        layout.addWidget(outputFileNameLabel,8,0)
        layout.addWidget(self.outputFileNameLineEdit,9,0,1,2)
        layout.addWidget(outputFileSelectionBtn,9,3,1,1)
        layout.addWidget(archiveCheckBox)
        self.setLayout(layout)

    def alterMode(self):
        if self.runMode == 1:
            self.runMode = 2
            self.inputFile1NameLineEdit.clear()
            self.inputFile2NameLineEdit.clear()
            self.inputFile3NameLineEdit.clear()
            self.outputFileNameLineEdit.clear()
            self.archive = False
        else:
            self.runMode = 1
            self.inputFile1NameLineEdit.setText(os.path.normpath(self.inf1PH))
            self.inputFile2NameLineEdit.setText(os.path.normpath(self.inf2PH))
            self.inputFile2NameLineEdit.setText(os.path.normpath(self.inf3PH))
            self.outputFileNameLineEdit.setText(os.path.normpath(self.outfPH))
            self.archive = True

    def alterArchive(self):
        if self.archive:
            self.archive = False
        else:
            self.archive = True

    def selectAttacheFile(self):
        self.inf1PH_C = QtWidgets.QFileDialog.getOpenFileName(self, 'Choose Attache Data File..','',"Comma Separated Values File (*.csv);;Attache Data File (*.INP);;All Files (*)")
        self.inputFile1NameLineEdit.setText(os.path.normpath(self.inf1PH_C[0]))

    def selectMileageFile(self):
        self.inf2PH_C = QtWidgets.QFileDialog.getOpenFileName(self, 'Choose Mileage Data File..','',"Comma Separated Values File (*.csv);;All Files (*)")
        self.inputFile2NameLineEdit.setText(os.path.normpath(self.inf2PH_C[0]))

    def selectDCWRatesFile(self):
        self.inf3PH_C = QtWidgets.QFileDialog.getOpenFileName(self, 'Choose DCW Rates File..','',"Excel File (*.xlsx);;All Files (*)")
        self.inputFile3NameLineEdit.setText(os.path.normpath(self.inf3PH_C[0]))

    def selectOutputDirc(self):
        self.outfPH_C = QtWidgets.QFileDialog.getExistingDirectory(self, 'Choose Saving Directory..','', QtWidgets.QFileDialog.ShowDirsOnly)
        self.outputFileNameLineEdit.setText(os.path.normpath(self.outfPH_C))
        #self.outputFileNameLineEdit.setText((self.outfPH_C))

    def saveFile(self):
        AttacheFileName = self.field("AttacheFileName")
        OutputFileName = self.field("OutputFileName")
        with open("Output.txt", "w") as text_file:
            text_file.write(OutputFileName)
            text_file.write("\n")
            text_file.write("The Operation Mode is: {}".format(self.runMode))
            text_file.write("\n")
            if (not(OutputFileName == ".")) and (OutputFileName):
                text_file.write("Field is not empty!")
            else:
                text_file.write("Field is empty!")



if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    wizard = myApp()
    wizard.show()
    sys.exit(app.exec_())

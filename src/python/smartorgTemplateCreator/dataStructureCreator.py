from openpyxl import load_workbook
import json


class DataStructureJsonGenerator(object):
    def __init__(self, excelFileName, templateName=None):
        wb = load_workbook(filename=excelFileName)
        self.wb = wb
        self.model = wb['Model']
        
        if templateName == None:
            self.templateName = self.extractTemplateNameFrom(excelFileName)
        else:
            self.templateName = templateName
        self.baseTemplate = {"Inputs": [], "Outputs": [], "Description": "An Ulu Template", "ID": self.templateName, "ExcelFile": "%s_template.xlsm" % (self.templateName)}
        self.baseTemplate["Outputs"] = [{"Units":"million dollars","CellLink":"Model!TotalTAM","Type":"SCALAR","Display":"Total TAM","Key":"TotalTAM"},{"Units":"Probability","CellLink":"Model!EarlySuccess","Type":"SCALAR","Display":"Chance of Early Success","Key":"EarlySuccess"},{"Units":"Probability","CellLink":"Model!CrossChasm","Type":"SCALAR","Display":"Chance of Crossing Chasm|Early Success","Key":"CrossChasm"},{"Units":"Probability","CellLink":"Model!MainstreamSuccess","Type":"SCALAR","Display":"Chance of Mainstream Success|Early & Cross Chasm Success","Key":"MainstreamSuccess"},{"Units": "probability", "CellLink": "Model!pCrossChasm_total","Type": "SCALAR","Display": "Probability of Cross Chasm Success","Key": "pCrossChasm_total"},{"Units": "probability","CellLink": "Model!pMassMkt_total","Type": "SCALAR","Display": "Probability of Mass Market Success","Key": "pMassMkt_total"},{"Units":"multiple","CellLink":"Model!TotalPWMOIC_SweetSpot","Type":"SCALAR","Display":"Total PWMOIC (Sweet Spot)","Key":"TotalPWMOC_SweetSpot"},{"Units":"multiple","CellLink":"Model!TotalPWMOIC_TotalTAM","Type":"SCALAR","Display":"Total PWMOIC (Total TAM)","Key":"TotalPWMOIC_TotalTAM"},{"Units":"million dollars","CellLink":"Model!SweetspotTAM","Type":"SCALAR","Display":"Sweet Spot TAM","Key":"SweetSpotTAM"},{"Units":"fraction","CellLink":"Model!pFailCrossChasm","Type":"SCALAR","Display":"Probability of Failing to Cross Chasm","Key":"pFailCrossChasm"},{"Units":"dollars","CellLink":"Model!CumeCostCrossChasm","Type":"SCALAR","Display":"Cumulative Cost of Crossing Chasm","Key":"CumeCostCrossChasm"},{"Units":"dollars","CellLink":"Model!CumeCostEarlyStage","Type":"SCALAR","Display":"Cumulative Cost of Early Stage","Key":"CumeCostEarlyStage"},{"Units":"fraction","CellLink":"Model!pFailEarlyStage","Type":"SCALAR","Display":"Probability of Failing at Early Stage","Key":"pFailEarlyStage"}, {"Units": "multiple","CellLink": "Model!TotalPWMOIC_SweetSpot_Given_CrossChasmSuccess","Type": "SCALAR","Display": "PWMOIC_SweetSpot|Cross Chasm Success","Key": "TotalPWMOIC_SweetSpot_Given_CrossChasmSuccess"},{"Units": "multiple","CellLink": "Model!TotalPWMOIC_TotalTAM_Given_CrossChasmSuccess","Type": "SCALAR","Display": "PWMOIC_TotalTAM|Cross Chasm Success","Key": "TotalPWMOIC_TotalTAM_Given_CrossChasmSuccess"}, {"Units": "multiple","CellLink": "Model!TotalPWMOIC_SweetSpot_Given_MarketSuccess","Type": "SCALAR","Display": "PWMOIC_SweetSpot|Market Success","Key": "TotalPWMOIC_SweetSpot_Given_MarketSuccess"}]
        self.baseTemplate["Inputs"] = [{"Description":"How much is being invested?","Val":0,"Constraint":"double","CellLink":"Model!Investment","Key":"Investment","Units":"million dollars","Type":"SCALAR","Display":"Investment Amount"},{"Description":"What is the pre-money valuation?","Val":[0,0,0],"Constraint":"double","CellLink":"Model!PreMoneyValue","Key":"PreMoneyValue","Units":"million dollars","Type":"DISTRIBUTION","Display":"Pre-money valuation"},{"Description":"What is the initial round size?","Val":[0,0,0],"Constraint":"double","CellLink":"Model!InitialRoundSize","Key":"InitialRoundSize","Units":"million dollars","Type":"DISTRIBUTION","Display":"Investment Size"},{"Description":"What is the fraction of dilution?","Val":[0,0,0],"Constraint":"double","CellLink":"Model!Dilution","Key":"Dilution","Units":"fraction","Type":"DISTRIBUTION","Display":"Dilution"},{"Description":"What is the fraction of niche dilution? Niche dilution represents how much dilution an Ulu investment would experience in this scenario. It is less than the mass market scenario as there are no late stage investment rounds or IPO.","Val":[0,0,0],"Constraint":"double","CellLink":"Model!NicheDilution","Key":"NicheDilution","Units":"fraction","Type":"DISTRIBUTION","Display":"Niche Dilution"},{"Description":"What is the chance that you will get the best wins (king)?","Val":0,"Constraint":"double","CellLink":"Model!KingChance","Key":"KingChance","Units":"fraction","Type":"SCALAR","Display":"King/Gorilla chance"},{"Description":"What is the chance that you will get modest wins (monkey)?","Val":0,"Constraint":"double","CellLink":"Model!PrinceChance","Key":"PrinceChance","Units":"fraction","Type":"SCALAR","Display":"Prince/Monkey chance"},{"Description":"What is the market share of the king?","Val":[0,0,0],"Constraint":"double","CellLink":"Model!KingMktShare","Key":"KingMktShare","Units":"fraction","Type":"DISTRIBUTION","Display":"King Market Share"},{"Description":"What is the market share of the prince?","Val":[0,0,0],"Constraint":"double","CellLink":"Model!PrinceMktShare","Key":"PrinceMktShare","Units":"fraction","Type":"DISTRIBUTION","Display":"Prince Market Share"},{"Description":"What is the market share of the serf?","Val":[0,0,0],"Constraint":"double","CellLink":"Model!SerfMktShare","Key":"SerfMktShare","Units":"fraction","Type":"DISTRIBUTION","Display":"Serf Market Share"},{"Description":"What is the niche size (as a fraction of total TAM)? Niche Size represents how large the niche market is as a percent of the overall market.","Val":[0,0,0],"Constraint":"double","CellLink":"Model!NicheSize","Key":"NicheSize","Units":"fraction","Type":"DISTRIBUTION","Display":"Niche Size"},{"Description":"Niche Market Share is your share of the niche market in this strategy. There are usually many more players in niches than in the mass market.","Val":[0,0,0],"Constraint":"double","CellLink":"Model!NicheMarketShare","Key":"NicheMarketShare","Units":"fraction","Type":"DISTRIBUTION","Display":"Niche Market Share"},{"Description":"What is the exit multiple?","Val":[0,0,0],"Constraint":"double","CellLink":"Model!SalesMult","Key":"SalesMult","Units":"number","Type":"DISTRIBUTION","Display":"Exit Multiple (Sales)"},{"Description":"Niche Exit Multiple Discount - in the niche scenario, the exit multiple is less (i.e. discounted) versus the exit multiple of a mass market player. The growth prospects of a company stuck in a niche are much more limited and this is reflected in the exit multiples these companies are able to command.","Val":[0,0,0],"Constraint":"double","CellLink":"Model!NicheMultDisc","Key":"NicheMultDisc","Units":"fraction","Type":"DISTRIBUTION","Display":"Niche Exit Mult Disc"},{"Val":0,"Constraint":"double","CellLink":"Model!earlyStage_market","Key":"earlyStage_market","Units":"fraction","Type":"SCALAR","Display":"Chance of early stage market success"},{"Val":0,"Constraint":"double","CellLink":"Model!earlyStage_product","Key":"earlyStage_product","Units":"fraction","Type":"SCALAR","Display":"Chance of early stage product success"},{"Val":0,"Constraint":"double","CellLink":"Model!earlyStage_team","Key":"earlyStage_team","Units":"fraction","Type":"SCALAR","Display":"Chance of early stage team success"},{"Val":0,"Constraint":"double","CellLink":"Model!earlyStage_financial","Key":"earlyStage_financial","Units":"fraction","Type":"SCALAR","Display":"Chance of early stage financial success"},{"Val":0,"Constraint":"double","CellLink":"Model!crossChasm_market","Key":"crossChasm_market","Units":"fraction","Type":"SCALAR","Display":"Chance of cross chasm market success"},{"Val":0,"Constraint":"double","CellLink":"Model!crossChasm_product","Key":"crossChasm_product","Units":"fraction","Type":"SCALAR","Display":"Chance of cross chasm product success"},{"Val":0,"Constraint":"double","CellLink":"Model!crossChasm_team","Key":"crossChasm_team","Units":"fraction","Type":"SCALAR","Display":"Chance of cross chasm team success"},{"Val":0,"Constraint":"double","CellLink":"Model!crossChasm_financial","Key":"crossChasm_financial","Units":"fraction","Type":"SCALAR","Display":"Chance of cross chasm financial success"},{"Val":0,"Constraint":"double","CellLink":"Model!massMarket_market","Key":"massMarket_market","Units":"fraction","Type":"SCALAR","Display":"Chance of market success in mass market"},{"Val":0,"Constraint":"double","CellLink":"Model!massMarket_product","Key":"massMarket_product","Units":"fraction","Type":"SCALAR","Display":"Chance of product success in mass market"},{"Val":0,"Constraint":"double","CellLink":"Model!massMarket_team","Key":"massMarket_team","Units":"fraction","Type":"SCALAR","Display":"Chance of team success in mass market"},{"Val":0,"Constraint":"double","CellLink":"Model!massMarket_financial","Key":"massMarket_financial","Units":"fraction","Type":"SCALAR","Display":"Chance of financial success in mass market"},  {"Val": "2024","Constraint": "year","CellLink": "Model!projectedYear","Key": "projectedYear","Inherited": True,"Units": "year","Type": "SCALAR","Display": "Projected Year"}]

    def extractTemplateNameFrom(self, excelFileName):
        nameWithoutExt = excelFileName[0:excelFileName.find('.')]
        posOfLastSlash = nameWithoutExt.rfind('/')
        return nameWithoutExt[posOfLastSlash+1:len(nameWithoutExt)]

    def addr(self, sheetName, cell):
        return '%s!%s' %(sheetName,cell.coordinate)

    def match(self, inputKey):
        inputs = self.baseTemplate["Inputs"]
        answer = None
        for input in inputs:
            if input['Key'] == inputKey:
                answer = input
        return answer

    def adjustUnits(self, units):
        if units == "%":
            return "fraction"
        else:
            return units

    def adjust(self, units):
        if '%' in units:
            return units.replace("%","fraction")
        else:
            return units

    def produceJSON(self):
        inputs = self.baseTemplate["Inputs"]
        address = list(self.wb.defined_names['inputTable'].destinations)
        numRow, cell, sheetname = self.startCellNumRowsAndSheetNameIn(address)
        for i in range(0,numRow):
            unitCell = cell.offset(column=2,row=i)
            if unitCell.value != None:
                descriptionCell = unitCell.offset(column=-1)
                valueCell = descriptionCell.offset(column=12)
                linkCell = valueCell.offset(column=1)
                results = { 'title': self.adjust(descriptionCell.value), 'units': self.adjust(unitCell.value), 'ref': self.addr(sheetname,valueCell), 'key': linkCell.value}
                lowCell = unitCell.offset(column=2)
                if lowCell.value != None:
                    baseCell = lowCell.offset(column=1)
                    highCell = lowCell.offset(column=2)
                    results['downloadLink'] = [self.addr(sheetname,lowCell), self.addr(sheetname,baseCell), self.addr(sheetname,highCell)]
                input = self.match(results['key'])
                if input == None:
                    inputs.append(self.makeJSON(results))
                else:
                    if 'downloadLink' in results:
                        input['DownloadLink'] = results['downloadLink']
                        input['CellLink'] = results['ref']

        return self.baseTemplate

    def startCellNumRowsAndSheetNameIn(self, address):
        for sheetname, cellAddress in address:
            cellAddress = cellAddress.replace('$','')
        cellComponents = cellAddress.split(':')
        startCell = cellComponents[0]
        endCell = cellComponents[1]
        cell = self.model[startCell]
        sCell = cell.offset(column=2)
        eCell = self.model[endCell]
        numRow = eCell.row - sCell.row + 1
        return numRow, cell, sheetname

    def makeKeyFrom(self, text):
        return text.replace(" ", "")

    def makeJSON(self,inputObj):
        if 'downloadLink' in inputObj:
            return {
                'Display': inputObj['title'],
                'Units': inputObj['units'],
                'CellLink': inputObj['ref'],
                'DownloadLink': inputObj['downloadLink'],
                'Val': [0,0,0],
                'Constraint': 'double',
                'Type': 'DISTRIBUTION',
                'Key': self.makeKeyFrom(inputObj['key'])
            }
        else:
            return {
                'Display': inputObj['title'],
                'Units': inputObj['units'],
                'Description': '0: Excluded, 1: Included',
                'CellLink': inputObj['ref'],
                'Val': 0,
                'Constraint': 'double',
                'Type': 'SCALAR',
                'Key': self.makeKeyFrom(inputObj['key'])
            }


dsGen = DataStructureJsonGenerator(excelFileName='/Users/somik/Ulu Ventures Dropbox/Somik Raha/Templates/MileAuto/MileAutoV2.xlsm')
print(dsGen.templateName)
template = dsGen.produceJSON()
print(json.dumps(template, indent=4))

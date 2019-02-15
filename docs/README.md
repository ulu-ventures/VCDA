# VCDA
Decision Analysis tools for Venture Capital

## Introduction
This project has been seeded with a spreadsheet model that has been developed by Clint Korver over a decade of Decision Analysis of a 100 seed-stage venture deals. The model is designed for enterprise company deals and implements two popular frameworks for enterprise-focused startups -- Crossing the Chasm and The Gorilla Game, both by Jeffrey Moore. 

The spreadsheet model contains special VBA software developed by Somik Raha to perform uncertainty analysis. If the basic structuring of the spreadsheet is followed, then any model within that structure can take advantage of the tools of uncertainty. Using the [VentureDeal model](https://github.com/ulu-ventures/VCDA/blob/master/spreadsheets/VentureDeal.xlsm), we will explain how the uncertainty analysis works. You can also learn about building the [TAM model](https://github.com/ulu-ventures/VCDA/blob/master/docs/TAM.md).

## Uncertainty Analysis
There are three key tables in the model that the VBA macros look for. They are:
* inputTable
* outputDefnTable
* SmartOrgInputTable

All three are required. 
![Annotated output ranges](https://github.com/ulu-ventures/VCDA/blob/master/docs/images/outputDefnTable.png)
![Annotated ranges](https://github.com/ulu-ventures/VCDA/blob/master/docs/images/annotationOfRanges.png)

Also, two sheets that should be kept intact are:
* Tornado
* Full Distribution

### inputTable
This contains the definitions of inputs. Please keep the same spacing format as you see in the VentureDeal template. The macro knows how to figure out how many inputs there are, and the low-base-high ranges for each. It also figures out if the input is a scalar or a distribution. 

### outputTable
This defines the output metrics of interest, and has the base case values pre-calculated. You will want to ensure that this is wired up correctly. You can add more outputs that you are interested in by ensuring they are in this table. The final distribution chart relies on "Target Mkt: PWMOIC given Cross Chasm Success", so if that gets moved around, you will want to check the Full Distribution sheet to ensure it is using the correct output metric.

### SmartOrgInputTable
This table is the one that drives the base case calculations. It is designed to be interoperable as a SmartOrg cloud-based template. You will notice that there are no formulae indexing on the "Index" column. And yet, if you change the index, the input in this table will change, as will the output metrics. A macro behind the scenes is taking care of this. Even if you are not deploying the model on the cloud, you will want to keep this table intact as the macro connected to this table is used when doing uncertainty calculations.

### The general flow of calculations
The steps of the uncertainty algorithm are as follows:
1. Build Tornado charts for each output
2. Take the top 5 bars in the Tornado and produce a distribution on each using a convolution algorithm
3. Summarize the distribution in the output metrics (Mean, Ten, Fifty, Ninety) next to each output
4. Produce the full distribution using the PWMOIC given Cross Chasm Success distribution and the lifestage probabilities

Clicking on the "Evaluate Uncertainty" button will complete 1, 2 and 3 and show the results. On the Tornado page, you will need to use the drop down to select the output metric of interest and click "Update Tornado" as show below.

![Updating Tornado](https://github.com/ulu-ventures/VCDA/blob/master/docs/images/updateTornado.png)

### Building the Tornado
Here is the algorithm for constructing a Tornado chart:
1. For each output:
  Vary each input, going with Low, Base High values.
2. For each variation, collect the corresponding outputs, sort and present as a Tornado.

Tornado digrams bring together two key ideas -- [leverage and uncertainty](https://smartorg.com/tornado-diagram-resolving-conflict-and-confusion-with-objectivity-and-evidence/). 

### Summarizing the Tornado with a Distribution
Once the Tornado is built, the macro pulls off the top 5 factors and builds a tree with each factor ranging from low to base to high. This operation is called convolution and is recursive in nature. Due to exponential time complexity, we limit it to the top 5 factors which more often than not account for 80-90% of the uncertainty. The convolution gives us a distribution that can be visualized as a Cumulative Distribution Function. The CDF chart is right below the Tornado chart for a given output.
![CDF](https://github.com/ulu-ventures/VCDA/blob/master/docs/images/tornadoCDF.png)

### Building the Full Distribution
We combine the lifestage uncertainties with the PWMOIC Distribution given Success to get to the Full Distribution. This is done by hardwiring the Full Distributiopn table (on the Full Distribution sheet) to the output PWMOIC Given Cross Chasm Success. Once wired up, the second table to the right makes the adjustment in cell S6 and S7, using pCrossChasm_total to calculate the probability of failure. The remaining probabilities are scaled appropriately. This produces a powerful visualization of the full uncertainty in the deal and lets VCs communicate to entrepreneurs the chance that they will not get any return whatsoever. When the long tail nature of the VC business is captured effectively, entrepreneurs no longer have the pressure to exaggerate their chances of success and can lean into the uncertainty. The Mean PWMOIC calculated is the one used for evaluation and tends to be higher than the Base PWMOIC.

![Full Distribution](https://github.com/ulu-ventures/VCDA/blob/master/docs/images/fullDistribution.png)




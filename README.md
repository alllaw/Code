preprocessing2.vba is designed to take Haringey_All.csv and DomesticBulkDataAB.xlsx combine the two data files and then process the results to allow them to be input into energyUsage.mat which uses
the SUSDEM modeling system. Haringey_All.csv comes from the output of M0.py using the UKMaps and UKBuildings shapefile. DomesticBulkDataAB.xlsx comes from the different EPC certificates. Haringey_All.csv
and DomesticBulkDataAB.xlsx are combined by formatting the addresses of the two different data files so that they are the same then comparing these addresses. If the addresses match then the data from
both Haringey_All.csv and DomesticBulkDataAB.xlsx are combined on the same row. It is often the case that there are multiple entries in DomesticBulkDataAB.xlsx for a single entry in Haringey_All.csv.
If this is the case all entries from DomesticBulkDataAB.xlsx that correspond to the entry in Haringey_All.csv are copied across and put in the rows below the initial Haringey_All.csv entry. After this
worksheet has been produced, the dates of the entries from DomesticBulkDataAB.xlsx are compared, the entries with the most recent dates are retained and all other entries are discared. If there are two
entries with the same, most recent date, as can be the case if one house is split into multiple flats with multiple EPC certificates, both EPC certificate entries are retained.

Once the file containing the combined entries from both Haringey_All.csv and DomesticBulkDataAB.xlsx name Haringey_Processed.xlsx has been produced it is subsequently processed to produce SUSDEMinput.xlsx.
SUSDEMinput.xlsx can then be used as an input to energyUsage.mat. When multiple DomesticBulkDataAB.xlsx entries for a single Haringey_All.csv enty are present, preprocessing2.vba takes the average of the
different DomesticBulkDataAB.xlsx entries. energyUsage.mat is concerned with showing the energy usage of the different building archetypes therfore entries into SUSDEMinput.xlsx that do not belong to an
archetype are deleted.

To run preprocessing2.vba, simply produce an Excel workbook containing one worksheet called Haringey_All which contains the data from Haringey_All.csv and one worksheet called DomesticBulkDataAB which
contains the data from DomesticBulkDataAB.xlsx. Copy this code into the visual basic editor and run it. preprocessing2.vba will produce a number of different worksheets for intermediate processes that
will be subsequently deleted at the end of the script.  

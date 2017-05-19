function [enduseDemand, annualDemand] = energyUsage()
%Takes data from SUSDEMinput and produces annual energy demand per m^2
ExTemp = xlsread('weatherInput','','A1:A12');
Irradiation = xlsread('irradiationInput','','A1:I12');

NoOfRooms = xlsread('SUSDEMinput3.xlsx','','E2:E1947');
NoOfStoreys = xlsread('SUSDEMinput3.xlsx','','F2:F1947');
WWR = xlsread('SUSDEMinput3.xlsx','','V2:V1947');
PercentageDoubleGlazing = xlsread('SUSDEMinput3.xlsx','','D2:D1947');
RoofInsulation = xlsread('SUSDEMinput3.xlsx','','AB2:AB1947');
PercentageLEL = ones(1946,1)*100;
HeatingSP = ones(1946,1)*20;
CoolingSP = ones(1946,1)*26;
COP = ones(1946,1)*0.8;
WaterTankInsulation = ones(1946,1)*0.25;
FractionHeated = ones(1946,1)*0.8;
BoilerEfficiency = ones(1946,1)*0.9;
CoolingCOP = ones(1946,1)*3;
NatVent = ones(1946,1)*0.75;
GlazingArea = ones(1946,1)*12;

%DwellingType = ones(1303,1); 
DwellingType = xlsread('SUSDEMinput3.xlsx','','A2:A1947');
DwellingPosition = xlsread('SUSDEMinput3.xlsx','','B2:B1947');
Orientation = ones(1946,1)*4;
%FloorConstruction = ones(1303,1); 
FloorConstruction = xlsread('SUSDEMinput3.xlsx','','U2:U1947');
ExternalWall1 = xlsread('SUSDEMinput3.xlsx','','S2:S1947');
InternalWall = ones(1946,1)*1.75;
DoorConstruction = ones(1946,1)*2;
ThermalMass = ones(1946,1)*2.6e5;
IlluminanceLevel = ones(1946,1)*90;
ShadingDevice = ones(1946,1)*1;
Infiltration = ones(1946,1)*2;
HeatingType1 = ones(1946,1)*1;
HeatingType2 = ones(1946,1)*2;
WaterHeating = ones(1946,1)*1;
SingleGlazing = ones(1946,1)*5.3; %make more accurate for different archetypes
DoubleGlazing = ones(1946,1)*2.1; %above
LELFactor = ones(1946,1)*0.25;
HouseholdNumber = ones(1946,1)*2;
CapitaConsumption = ones(1946,1)*2.3;
ExternalWall2 = xlsread('SUSDEMinput3.xlsx','','T2:T1947');
DwellingAge = xlsread('SUSDEMinput3.xlsx','','R2:R1947');
HouseholdConsumption = ones(1946,1)*7.5;

FloorAreas = xlsread('SUSDEMinput3.xlsx','','I2:I1947');
FloorHeights = xlsread('SUSDEMinput3.xlsx','','L2:L1947');
FloorPerimeters = xlsread('SUSDEMinput3.xlsx','','O2:O1947');

for n=1:1945 % where N is the number of dwellings
    
aa = NoOfRooms(n,1);
ab = NoOfStoreys(n,1);
ac = WWR(n,1);
ad = PercentageDoubleGlazing(n,1);
ae = RoofInsulation(n,1);
af = PercentageLEL(n,1);
ag = HeatingSP(n,1);
ah = CoolingSP(n,1);
ai = COP(n,1);
aj = WaterTankInsulation(n,1);
ak = FractionHeated(n,1);
al = BoilerEfficiency(n,1);
am = CoolingCOP(n,1);
an = NatVent(n,1);
ao = GlazingArea(n,1);

ba = DwellingType(n,1);
bb = DwellingPosition(n,1);
bc = Orientation(n,1);
bd = FloorConstruction(n,1);
be = ExternalWall1(n,1);
bf = InternalWall(n,1);
bg = DoorConstruction(n,1);
bh = ThermalMass(n,1);
bi = IlluminanceLevel(n,1);
bj = ShadingDevice(n,1);
bk = Infiltration(n,1);
bl = HeatingType1(n,1);
bm = HeatingType2(n,1);
bn = WaterHeating(n,1);
bo = SingleGlazing(n,1);
bp = DoubleGlazing(n,1);
bq = LELFactor(n,1);
br = HouseholdNumber(n,1);
bs = CapitaConsumption(n,1);
bt = ExternalWall2(n,1);
bu = DwellingAge(n,1);
bv = HouseholdConsumption(n,1);

ca = FloorAreas(n,1);
cb = FloorHeights(n,1);
cc = FloorPerimeters(n,1);

Inputs1 = [aa 1 ac ad ae af ag ah ai aj ak al am an ao];
%Inputs1 = [2 1 0.23 100 200 100 21 26 0.8 100 0.8 0.9 3 0.75 5];
Inputs2 = [ba bb bc bd be bf bg bh bi bj bk bl bm bn bo bp bq br bs bt bu bv];
%Inputs2 = [1 6 bc 2 0 1.75 2 2.6e5 bj 1 bl 1 bn bo 5.3 2.1 br bs 3 3 7.5];
%Inputs2 = [1 6 4 2 0 1.75 2 2.6e5 90 1 2 1 2 1 5.3 2.1 0.25 2 2.3 3 3 7.5];
Inputs3 = [ca cb cc];
%Inputs3 = [90 2.8 40];
Inputs4 = [ExTemp Irradiation];

%Calculates space heating, hot water, electrical, and cooling demand for each month (baseline - 2010)
[spaceHeating, DHW, Electricity] = ResidentialEnergyDemand(Inputs1, Inputs2, Inputs3, Inputs4); % for each dwelling

enduseDemand = [spaceHeating DHW Electricity]/ca; % to get per m2
annualDemand(n,1:3) = (sum(enduseDemand,1))'; % annual end use energy demand (kWh/m2*year) (6,1,1,:)
end
end

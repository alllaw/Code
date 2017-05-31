UKB = xlsread('energyinput2.xlsm','UKBnew','A2:B3345');
EPC = xlsread('energyinput2.xlsm','EPCscaled','A2:B2230');

TF1 = UKB(:,2)==0;
UKB(TF1,:) = [];
UKB = UKB(~any(isnan(UKB),2),:);

TF1 = EPC(:,2)==0;
EPC(TF1,:) = [];
EPC = EPC(~any(isnan(EPC),2),:);

UKBupn = UKB(:,1);
UKBenergy = UKB(:,2);
EPCfid = EPC(:,1);
EPCenergy = EPC(:,2);

sukb = std(UKBenergy);
mukb = mean(UKBenergy);
sepc = std(EPCenergy);
mepc = mean(EPCenergy);

figure
hist(UKBenergy,30)
title('Histogram Showing Energy Usage per Building From Model 2 Output un-edited')
ylabel('Frequency')
xlabel('Yearly Energy Consumption (Kwh)')
figure
hist(EPCenergy,30)
title('Histogram Showing Energy Usage per Building From EPC data un-edited')
ylabel('Frequency')
xlabel('Yearly Energy Consumption (Kwh)')

indices = abs(UKBenergy) > mukb + (1.5*sukb);
UKBenergy(indices) = [];
indices = abs(EPCenergy) > mepc + (1.5*sepc);
EPCenergy(indices) = [];

figure
hist(UKBenergy,30)
title('Histogram Showing Energy Usage per Building From Model 2 Output (\mu + 1.5\times\sigma removed) ')
ylabel('Frequency')
xlabel('Yearly Energy Consumption (Kwh)')
figure
hist(EPCenergy,30)
title('Histogram Showing Energy Usage per Building From EPC data (\mu + 1.5\times\sigma removed)')
ylabel('Frequency')
xlabel('Yearly Energy Consumption (Kwh)')

indices = abs(EPCenergy) > 60000;
EPCenergy(indices) = [];

figure
hist(EPCenergy,30)
title('Histogram Showing Energy Usage per Building From EPC data (same scale as Model 2)')
ylabel('Frequency')
xlabel('Yearly Energy Consumption (Kwh)')

TF1 = UKB(:,2) >= mukb + (1.5*sukb);
UKB(TF1,:) = [];
UKB = UKB(~any(isnan(UKB),2),:);

TF1 = EPC(:,2) >= 60000;
EPC(TF1,:) = [];
EPC = EPC(~any(isnan(EPC),2),:);








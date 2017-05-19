function [totalAnnualDemandArchetypes1a, totalAnnualDemandArchetypes1b, totalAnnualDemandArchetypes1c, ...
    totalAnnualDemandArchetypes2a, totalAnnualDemandArchetypes3a, totalAnnualDemandArchetypes4a, ...
    totalAnnualDemandArchetypes4b, totalAnnualDemandArchetypes4c, totalAnnualDemandArchetypes5a, ...
    totalAnnualDemandArchetypes5b] = archetypes()
[~, annualDemand] = energyUsage();
totalAnnualDemand = sum(annualDemand,2);
[~,archs,~] = xlsread('SUSDEMinput2.xlsx','','C2:C1947');
[mt,~] = size(totalAnnualDemand);
archs = archs(1:mt);
totalAnnualDemandArchetypes = [hex2dec(archs), totalAnnualDemand]; %Convert archertype classes to number for easy conversion
% 1a - 26, 1b - 27, 1c - 28
% 2a - 42
% 3a - 58
% 4a - 74, 4b - 75, 4c - 76
% 5a - 90, 5b - 91
%Find indices to elements in first column of totalAnnualDemandArchetypes that satisfy the equality
ind1 = totalAnnualDemandArchetypes(:,1) == 26;
ind2 = totalAnnualDemandArchetypes(:,1) == 27;
ind3 = totalAnnualDemandArchetypes(:,1) == 28;
ind4 = totalAnnualDemandArchetypes(:,1) == 42;
ind5 = totalAnnualDemandArchetypes(:,1) == 58;
ind6 = totalAnnualDemandArchetypes(:,1) == 74;
ind7 = totalAnnualDemandArchetypes(:,1) == 75;
ind8 = totalAnnualDemandArchetypes(:,1) == 76;
ind9 = totalAnnualDemandArchetypes(:,1) == 90;
ind10 = totalAnnualDemandArchetypes(:,1) == 91;
%Use the logical indices to index into totalAnnualDemandArchetypes to return required sub-matrices
totalAnnualDemandArchetypes1a = totalAnnualDemandArchetypes(ind1,:);
totalAnnualDemandArchetypes1b = totalAnnualDemandArchetypes(ind2,:);
totalAnnualDemandArchetypes1c = totalAnnualDemandArchetypes(ind3,:);
totalAnnualDemandArchetypes2a = totalAnnualDemandArchetypes(ind4,:);
totalAnnualDemandArchetypes3a = totalAnnualDemandArchetypes(ind5,:);
totalAnnualDemandArchetypes4a = totalAnnualDemandArchetypes(ind6,:);
totalAnnualDemandArchetypes4b = totalAnnualDemandArchetypes(ind7,:);
totalAnnualDemandArchetypes4c = totalAnnualDemandArchetypes(ind8,:);
totalAnnualDemandArchetypes5a = totalAnnualDemandArchetypes(ind9,:);
totalAnnualDemandArchetypes5b = totalAnnualDemandArchetypes(ind10,:);
end

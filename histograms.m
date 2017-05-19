[totalAnnualDemandArchetypes1a, totalAnnualDemandArchetypes1b, totalAnnualDemandArchetypes1c, ...
totalAnnualDemandArchetypes2a, totalAnnualDemandArchetypes3a, totalAnnualDemandArchetypes4a, ...
totalAnnualDemandArchetypes4b, totalAnnualDemandArchetypes4c, totalAnnualDemandArchetypes5a, ...
totalAnnualDemandArchetypes5b] = archetypes();

hist(totalAnnualDemandArchetypes1a(:,2))
figure
hist(totalAnnualDemandArchetypes1b(:,2))
figure
hist(totalAnnualDemandArchetypes1c(:,2))
figure
hist(totalAnnualDemandArchetypes2a(:,2))
figure
hist(totalAnnualDemandArchetypes3a(:,2))
figure
hist(totalAnnualDemandArchetypes4a(:,2))
figure
hist(totalAnnualDemandArchetypes4b(:,2))
figure
hist(totalAnnualDemandArchetypes4c(:,2))
figure
hist(totalAnnualDemandArchetypes5a(:,2))
figure
hist(totalAnnualDemandArchetypes5b(:,2))

% s3 = std(annualDemand);
% sT = std(totalAnnualDemand);
% m3 = mean(annualDemand);
% mT = mean(totalAnnualDemand);
% 
% indices = abs(totalAnnualDemand) < (mT - (2.5*sT));
% totalAnnualDemand(indices) = [];
% indices = abs(totalAnnualDemand) > (mT + (2.5*sT));
% totalAnnualDemand(indices) = [];
% hist(totalAnnualDemand,30)




[~,HaringeyPostCodes,~] = xlsread('postCodeAverages.xlsm','Sheet4','A1:A5');
[HaringeyEnergyUsageNum,HaringeyEnergyUsageText,HaringeyEnergyUsage] = xlsread('postCodeAverages.xlsm','Sheet1','A1:B560');

HPC = char(HaringeyPostCodes);
a = uint8(HPC);
b = double(a);
c = 0;
for r = 1:5
    h = b(:,r)*(100.^r);
    c = h + c;
end

HaringeyEnergyUsageText = char(HaringeyEnergyUsageText(:,1));
d = uint8(HaringeyEnergyUsageText);
e = double(d);
f = 0;
for r = 1:5
    h = e(:,r)*(100.^r);
    f = h + f;
end
g = [f(1:550,1), HaringeyEnergyUsageNum(1:550,1)];

for q = 1:5
    ind = g(:,1) == c(q,1);
    dum = g(ind,:);
    x = dum(:,2);
    index = x>0;
    n = sum(index);
    n(n==0)=NaN;
    M(q,1) = sum(x.*index)./n;
end

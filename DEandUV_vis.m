clear all; close all 
m = xlsread('PureICBAScoutresult.xlsx', 'Sheet1');
n = xlsread('PureICBAScoutresult.xlsx', 'Sheet2');
o = xlsread('PureICBAScoutresult.xlsx', 'Sheet3');
figure
[ax,h1,h2]=plotyy(o(:,1),o(:,3)*100,o(:,1),o(:,5)*100)

set(h1,'color','k')
set(h2,'linestyle','--','color','k')
set(ax(1),'XLim',[300,800],'YLim',[-10 100],'YTickmode','auto','Box','off','Fontsize',14,'Ycolor','k')
set(ax(2),'XLim',[300,800],'YLim',[0 60],'YTickmode','auto','Fontsize',14,'Ycolor','k')
line(o(:,1),o(:,2)*100,'Color', 'r', 'LineStyle', '-.', 'parent',ax(1))
line(downsample(o(:,1),13),downsample(o(:,4)*100,13),'Color', 'r', 'LineStyle', '--','parent',ax(2))
set(gca,'fontsize',14)
xlabel('Wavelength (nm)')
ylabel(ax(1),'Transmittance (%)','color','k')
ylabel(ax(2),'Reflectance (%)','color','k')
% Create textarrow
annotation('textarrow',[0.706231884057971 0.839304347826087],...
    [0.236835051546392 0.237835051546392],'TextEdgeColor','none');
% Create arrow
annotation('arrow',[0.332753623188406 0.243478260869565],...
    [0.535082474226804 0.534020618556701]);

w = m(:,1);
 w1 = n(:,1);
r = m(:,2);
 r1 = n(:,2);
i = m(:,3);
 i1 = n(:,3);


figure
[ax,h1,h2]=plotyy(w,r,w,i)

set(h1,'color','k')
set(h2,'linestyle','--','color','k')
set(ax(1),'XLim',[300,800],'YLim',[0 5],'YTickmode','auto','Box','off','Fontsize',14,'Ycolor','k')
set(ax(2),'XLim',[300,800],'YLim',[0 3],'YTickmode','auto','Fontsize',14,'Ycolor','k')
line(w1,r1,'Color', 'r', 'LineStyle', '--', 'parent',ax(1))
line(downsample(w1,13),downsample(i1,13),'Color', 'r', 'LineStyle', '--','parent',ax(2))
set(gca,'fontsize',14)
xlabel('Wavelength (nm)')
ylabel(ax(1),'Real Part','color','k')
ylabel(ax(2),'Imaginary Part','color','k')
title ('ICBA Dielectric Function')
% Create textarrow
annotation('textarrow',[0.306231884057971 0.209304347826087],...
    [0.701835051546392 0.70135051546392],'TextEdgeColor','none');
% Create arrow
annotation('arrow',[0.612753623188406 0.733478260869565],...
    [0.185082474226804 0.184020618556701]);




















% hold on
% plot(w,r,'-k',w,i,'-.k')
% plot(w1,r1,'--r',downsample(w1,3),downsample(i1,3),'.r','linewidth',2)
% axis([300 800 0 5])
% set(gca,'fontsize',14)
% %legend('Modeled Real','Modeled Imag.','Palik Real','Palik Imag.','Location','southwest')
% xlabel('Wavelength (nm)','fontsize',14)
% ylabel('Dielectric function','fontsize',14)
% title('P3HT Dielectric Function','fontsize',14)
% 
% norm=1/(length(o(:,1)));
% rmsr=(norm*sum(((o(:,2)-o(:,3)).^2)))
% rmsi=(norm*sum(((o(:,5)-o(:,4)).^2)))
% rmsr+rmsi


% 
% [ax,h1,h2]=plotyy(w,r,w,i)
% [ax2,h12,h22]=plotyy(w1,r1,w1,i1)
% 
% 
% set(h1,'color','k')
% set(h2,'linestyle','--','color','k')
% set(ax(1),'XLim',[300,800],'YLim',[0 4.5],'YTickmode','auto','Box','off','Fontsize',14,'Ycolor','k')
% set(ax(2),'XLim',[300,800],'YLim',[0 3],'YTickmode','auto','Fontsize',14,'Ycolor','k')
%  %line(o(:,1),o(:,2),'Color', 'r', 'LineStyle', '-.', 'parent',ax(1))
% % line(downsample(o(:,1),13),downsample(o(:,4),13),'Color', 'r', 'LineStyle', 'none', 'Marker', '.','parent',ax(2))
% 
% set(gca,'fontsize',14)
% xlabel('Wavelength (nm)')
% ylabel(ax(1),'Real Part','color','k')
% ylabel(ax(2),'Imaginary part','color','k')
% title('Complex Dielectric function');
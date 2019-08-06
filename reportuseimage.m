clc;%excel plot
clear all;%
close all;
openfile = 'C:\Users\lian\Desktop\march.xlsm';
[data, txt] = xlsread(openfile, 1);
[r,c]=size(txt);
windboundaryt = 35;
tiltboundaryt = 350;
windboundaryb = 30;
tiltboundaryb = 150;
windinterval = 5;
tiltinterval = 25;
for i = 84%1:r%i+13 = JD%wind event
    P23openname = cell2mat(txt(i,1));
    P24openname = cell2mat(txt(i,2));
    windopenname = cell2mat(txt(i,3));
    P23file = ['C:\Users\lian\Desktop\1\' P23openname];      
    P24file = ['C:\Users\lian\Desktop\1\' P24openname];  
    windfile = ['C:\Users\lian\Desktop\1\' windopenname];      
    year = str2double(P23openname(1:4));
    month = str2double(P23openname(5:6));
    day = str2double(P23openname(7:8));
    if year == 2008 %year2008- 29, year2009- 28, year2010- 28
        feb=29;
    else
        feb=28;
    end
    if month == 1
        jd = day;
    elseif month == 2
        jd = day+31;
    elseif month == 3
        jd = day+31+feb;
    elseif month == 4
        jd = day+31+feb+31;
    elseif month == 5
        jd = day+31+feb+31+30;
    end 
    
    if year == 2008
        top23equation=0.1693*jd-268.88;
        bottom23equation=3.2006*jd-189.07;
        top24equation=0.2687*jd-20.483;
        bottom24equation=0.227*jd-166.02;
    elseif year == 2009
        top23equation=0.9178*jd-255.94;
        bottom23equation=0.7499*jd+628.27;
        top24equation=0.6871*jd-68.621;
        bottom24equation=0.6301*jd-214.65;
    elseif year == 2010
        top23equation=1.9732*jd-387.18;
        bottom23equation=1.4557*jd+967.91;
        top24equation=1.8641*jd-147.8;
        bottom24equation=0.848*jd-219.83;      
    end
    
    P23tiltdata = xlsread(P23file, 4);  
    P24tiltdata = xlsread(P24file, 4);      
    winddata = xlsread(windfile, 4);      
    [P23datar,P23datac]=size(P23tiltdata);  
    [P24datar,P24datac]=size(P24tiltdata);  
    
    P23tiltjd= (P23tiltdata(:,1)-jd).*24;    
    P23toptilt = P23tiltdata(:,2)-top23equation+130-11-2+7+10-5;
    P23bottomtilt = P23tiltdata(:,3)-bottom23equation-15+2-5+4+0.5-9+5+3.9;
   
    P24tiltjd= (P24tiltdata(:,1)-jd).*24;
    P24toptilt = P24tiltdata(:,2)-top24equation+108-3+9;
    P24bottomtilt = P24tiltdata(:,3)-bottom24equation-5+4-5+3;
          
    windjd = (winddata(:,1)-jd).*24;
    if year ==2008
        windsd = winddata(:,2); %2008
    elseif year == 2009
        windsd = winddata(:,3); %2009 
    elseif year == 2010
        windsd = winddata(:,3); %2009 
    end          
    for j = 1:P23datar
        if P23toptilt(j) == max(P23toptilt)
            P23tmaxpoint = P23tiltjd(j);           
        elseif P23toptilt(j) == min(P23toptilt)
            P23tminpoint = P23tiltjd(j);           
        end
        if P23bottomtilt(j) == max(P23bottomtilt)
            P23bmaxpoint = P23tiltjd(j);          
        elseif P23bottomtilt(j) == min(P23bottomtilt)
            P23bminpoint = P23tiltjd(j);          
        end      
    end   
    for j = 1:P24datar
        if P24toptilt(j) == max(P24toptilt)
            P24tmaxpoint = P24tiltjd(j);           
        elseif P24toptilt(j) == min(P24toptilt)
            P24tminpoint = P24tiltjd(j);           
        end
        if P24bottomtilt(j) == max(P24bottomtilt)
            P24bmaxpoint = P24tiltjd(j);          
        elseif P24bottomtilt(j) == min(P24bottomtilt)
            P24bminpoint = P24tiltjd(j);          
        end      
    end   

    figure(1)
    subplot(2,1,1)
    [P23tophaxes, P23toptiltplot,P23twindplot] = plotyy(P23tiltjd,P23toptilt, windjd,-windsd);
    text(P23tmaxpoint,max(P23toptilt),['\leftarrow ' num2str(max(P23toptilt))],...
        'HorizontalAlignment','left')    
    text(P23tminpoint,min(P23toptilt),[num2str(min(P23toptilt)) '\rightarrow'],...
        'HorizontalAlignment','right')         
    title([P24openname(1:8) '(JD' num2str(jd) ')' ' Upper Tiltmeter'])   
    axes(P23tophaxes(1))
    xlim([0 24])      
    set(gca,'XTick',0:1:24)
    ylabel('P23 Tilt (\mu rad)')
    ylim([-50 tiltboundaryt])      
    set(gca,'YTick',-50:tiltinterval:tiltboundaryt)   
    axes(P23tophaxes(2))
    xlim([0 24])      
    set(gca,'XTick',0:1:24)
    ylabel('Wind Speed (m/s)')
    ylim([-5 windboundaryt])          
    set(gca,'YTick',-5:windinterval:windboundaryt) 
    set(P23twindplot,'LineStyle',':','LineWidth',2)
    Grid on 
    
    subplot(2,1,2)
    [P24tophaxes, P24toptiltplot,P24twindplot] = plotyy(P24tiltjd,P24toptilt, windjd,-windsd);        
    text(P24tmaxpoint,max(P24toptilt),['\leftarrow ' num2str(max(P24toptilt))],...
        'HorizontalAlignment','left')    
    text(P24tminpoint,min(P24toptilt),[num2str(min(P24toptilt)) '\rightarrow'],...
        'HorizontalAlignment','right')   
        xlabel('Day Time (hour)'); 
     
    axes(P24tophaxes(1))
    xlim([0 24])      
    set(gca,'XTick',0:1:24)
    ylabel('P24 Tilt (\mu rad)') 
    ylim([-50 tiltboundaryt])      
    set(gca,'YTick',-50:tiltinterval:tiltboundaryt)   
    axes(P24tophaxes(2))
    xlim([0 24])      
    set(gca,'XTick',0:1:24)
    ylabel('Wind Speed (m/s)')
    ylim([-5 windboundaryt])          
    set(gca,'YTick',-5:windinterval:windboundaryt) 
    set(P24twindplot,'LineStyle',':','LineWidth',2)
    Grid on 
   
    savefile = ['C:\Users\lian\Desktop\2\' P23openname(1:11) 'Toptilt'];
    saveas(P23toptiltplot, savefile , 'jpg')     
    %close all
          
    figure(2)
    subplot(2,1,1)
        [P23bottomhaxes, P23bottomtiltplot,P23bwindplot] = plotyy(P23tiltjd,P23bottomtilt, windjd,-windsd);
    text(P23bmaxpoint,max(P23bottomtilt),['\leftarrow ' num2str(max(P23bottomtilt))],...
        'HorizontalAlignment','left')    
    text(P23bminpoint,min(P23bottomtilt),['\leftarrow ' num2str(min(P23bottomtilt))],...
        'HorizontalAlignment','left')   
    title([P24openname(1:8) '(JD' num2str(jd) ')' ' Lower Tiltmeter'])
    %xlabel('Day Time (hour)');   
    axes(P23bottomhaxes(1))
    xlim([0 24])      
    set(gca,'XTick',0:1:24)
    ylabel('P23 Tilt (\mu rad)')
    ylim([-25 tiltboundaryb])      
    set(gca,'YTick',-25:25:tiltboundaryb)
    axes(P23bottomhaxes(2))
    xlim([0 24])      
    set(gca,'XTick',0:1:24)
    ylabel('Wind Speed (m/s)')
    ylim([-5 windboundaryb])          
    set(gca,'YTick',-5:windinterval:windboundaryb) 
    set(P23bwindplot,'LineStyle',':','LineWidth',2)
    Grid on   
    
    subplot(2,1,2)
    [P24bottomhaxes, P24bottomtiltplot,P24bwindplot] = plotyy(P24tiltjd,P24bottomtilt, windjd,-windsd);        
    text(P24bmaxpoint,max(P24bottomtilt),['\leftarrow ' num2str(max(P24bottomtilt))],...
        'HorizontalAlignment','left')    
    text(P24bminpoint,min(P24bottomtilt),['\leftarrow ' num2str(min(P24bottomtilt))],...
        'HorizontalAlignment','left')   
%     text(P24bminpoint,min(P24bottomtilt),[num2str(min(P24bottomtilt)) '\rightarrow'],...
%         'HorizontalAlignment','right')        
    xlabel('Day Time (hour)');      
    axes(P24bottomhaxes(1))
    xlim([0 24])      
    set(gca,'XTick',0:1:24)
    ylabel('P24 Tilt (\mu rad)')
    ylim([-25 tiltboundaryb])      
    set(gca,'YTick',-25:25:tiltboundaryb)
    axes(P24bottomhaxes(2))
    xlim([0 24])      
    set(gca,'XTick',0:1:24)
    ylabel('Wind Speed (m/s)')
    ylim([-5 windboundaryb])          
    set(gca,'YTick',-5:windinterval:windboundaryb) 
    set(P24bwindplot,'LineStyle',':','LineWidth',2)
    Grid on 
    
    savefile = ['C:\Users\lian\Desktop\2\' P24openname(1:11) 'Bottomtilt'];
    saveas(P23bottomtiltplot, savefile , 'jpg')     
    %close all
         
end    


% 
% clc;%excel plot
% clear all;%
% close all;
% openfile = 'C:\Users\lian\Desktop\march.xlsm';
% [data, txt] = xlsread(openfile, 1);
% [r,c]=size(txt);
% windboundaryt = 20;
% tiltboundaryt = 300;
% windboundaryb = 20;
% tiltboundaryb = 120;
% windinterval = 5;
% tiltinterval = 30;
% for i = 51%1:r%i+13 = JD%ice event
%     P23openname = cell2mat(txt(i,1));
%     P24openname = cell2mat(txt(i,2));
%     windopenname = cell2mat(txt(i,3));
%     P23file = ['C:\Users\lian\Desktop\1\' P23openname];      
%     P24file = ['C:\Users\lian\Desktop\1\' P24openname];  
%     windfile = ['C:\Users\lian\Desktop\1\' windopenname];      
%     year = str2double(P23openname(1:4));
%     month = str2double(P23openname(5:6));
%     day = str2double(P23openname(7:8));
%     if year == 2008 %year2008- 29, year2009- 28, year2010- 28
%         feb=29;
%     else
%         feb=28;
%     end
%     if month == 1
%         jd = day;
%     elseif month == 2
%         jd = day+31;
%     elseif month == 3
%         jd = day+31+feb;
%     elseif month == 4
%         jd = day+31+feb+31;
%     elseif month == 5
%         jd = day+31+feb+31+30;
%     end 
%     
%     if year == 2008
%         top23equation=0.1693*jd-268.88;
%         bottom23equation=3.2006*jd-189.07;
%         top24equation=0.2687*jd-20.483;
%         bottom24equation=0.227*jd-166.02;
%     elseif year == 2009
%         top23equation=0.9178*jd-255.94;
%         bottom23equation=0.7499*jd+628.27;
%         top24equation=0.6871*jd-68.621;
%         bottom24equation=0.6301*jd-214.65;
%     elseif year == 2010
%         top23equation=1.9732*jd-387.18;
%         bottom23equation=1.4557*jd+967.91;
%         top24equation=1.8641*jd-147.8;
%         bottom24equation=0.848*jd-219.83;      
%     end
%     
%     P23tiltdata = xlsread(P23file, 4);  
%     P24tiltdata = xlsread(P24file, 4);      
%     winddata = xlsread(windfile, 4);      
%     [P23datar,P23datac]=size(P23tiltdata);  
%     [P24datar,P24datac]=size(P24tiltdata);  
%     
%     P23tiltjd= (P23tiltdata(:,1)-jd).*24;    
%     P23toptilt = P23tiltdata(:,2)-top23equation-82-2;
%     P23bottomtilt = P23tiltdata(:,3)-bottom23equation+28-1.5;
%    
%     P24tiltjd= (P24tiltdata(:,1)-jd).*24;
%     P24toptilt = P24tiltdata(:,2)-top24equation-50;
%     P24bottomtilt = P24tiltdata(:,3)-bottom24equation+9-2;
%           
%     windjd = (winddata(:,1)-jd).*24;
%     if year ==2008
%         windsd = winddata(:,2); %2008
%     elseif year == 2009
%         windsd = winddata(:,3); %2009 
%     elseif year == 2010
%         windsd = winddata(:,3); %2009 
%     end          
%     for j = 1:P23datar
%         if P23toptilt(j) == max(P23toptilt)
%             P23tmaxpoint = P23tiltjd(j);           
%         elseif P23toptilt(j) == min(P23toptilt)
%             P23tminpoint = P23tiltjd(j);           
%         end
%         if P23bottomtilt(j) == max(P23bottomtilt)
%             P23bmaxpoint = P23tiltjd(j);          
%         elseif P23bottomtilt(j) == min(P23bottomtilt)
%             P23bminpoint = P23tiltjd(j);          
%         end      
%     end   
%     for j = 1:P24datar
%         if P24toptilt(j) == max(P24toptilt)
%             P24tmaxpoint = P24tiltjd(j);           
%         elseif P24toptilt(j) == min(P24toptilt)
%             P24tminpoint = P24tiltjd(j);           
%         end
%         if P24bottomtilt(j) == max(P24bottomtilt)
%             P24bmaxpoint = P24tiltjd(j);          
%         elseif P24bottomtilt(j) == min(P24bottomtilt)
%             P24bminpoint = P24tiltjd(j);          
%         end      
%     end   
% 
% %     figure(1)
% %     subplot(2,1,1)
% %     [P23tophaxes, P23toptiltplot,P23twindplot] = plotyy(P23tiltjd,P23toptilt, windjd,-windsd);
% %     text(P23tmaxpoint,max(P23toptilt),['\leftarrow ' num2str(max(P23toptilt))],...
% %         'HorizontalAlignment','left')    
% %     text(P23tminpoint,min(P23toptilt),[num2str(min(P23toptilt)) '\rightarrow'],...
% %         'HorizontalAlignment','right')         
% %     title(['Top Tilt ' P23openname(1:11) ' (JD' num2str(jd) ')'])     
% %     axes(P23tophaxes(1))
% %     xlim([0 24])      
% %     set(gca,'XTick',0:1:24)
% %     ylabel('Tilt (\mu rad)')
% %     ylim([-tiltboundaryt tiltboundaryt])      
% %     set(gca,'YTick',-tiltboundaryt:tiltinterval:tiltboundaryt)   
% %     axes(P23tophaxes(2))
% %     xlim([0 24])      
% %     set(gca,'XTick',0:1:24)
% %     ylabel('Wind Speed (m/s)')
% %     ylim([-windboundaryt windboundaryt])          
% %     set(gca,'YTick',-windboundaryt:windinterval:windboundaryt) 
% %     set(P23twindplot,'LineStyle',':','LineWidth',2)
% %     Grid on 
% %     
% %     subplot(2,1,2)
% %     [P24tophaxes, P24toptiltplot,P24twindplot] = plotyy(P24tiltjd,P24toptilt, windjd,-windsd);        
% %     text(P24tmaxpoint,max(P24toptilt),['\leftarrow ' num2str(max(P24toptilt))],...
% %         'HorizontalAlignment','left')    
% %     text(P24tminpoint,min(P24toptilt),[num2str(min(P24toptilt)) '\rightarrow'],...
% %         'HorizontalAlignment','right')   
% %         xlabel('Day Time (hour)'); 
% %      
% %     axes(P24tophaxes(1))
% %     xlim([0 24])      
% %     set(gca,'XTick',0:1:24)
% %     ylabel('Tilt (\mu rad)') 
% %     ylim([-tiltboundaryt tiltboundaryt])      
% %     set(gca,'YTick',-tiltboundaryt:tiltinterval:tiltboundaryt)   
% %     axes(P24tophaxes(2))
% %     xlim([0 24])      
% %     set(gca,'XTick',0:1:24)
% %     ylabel('Wind Speed (m/s)')
% %     ylim([-windboundaryt windboundaryt])          
% %     set(gca,'YTick',-windboundaryt:windinterval:windboundaryt) 
% %     set(P24twindplot,'LineStyle',':','LineWidth',2)
% %     Grid on 
% %    
% %     savefile = ['C:\Users\lian\Desktop\2\' P23openname(1:11) 'Toptilt'];
% %     saveas(P23toptiltplot, savefile , 'jpg')     
% %     %close all
%           
%     figure(2)
% %     subplot(2,1,1)
%         [P23bottomhaxes, P23bottomtiltplot,P23bwindplot] = plotyy(P23tiltjd,P23bottomtilt, windjd,-windsd);
% %     text(P23bmaxpoint,max(P23bottomtilt),['\leftarrow ' num2str(max(P23bottomtilt))],...
% %         'HorizontalAlignment','left')    
%     text(P23bminpoint,min(P23bottomtilt),[num2str(min(P23bottomtilt)) '\rightarrow'],...
%         'HorizontalAlignment','right')   
%     title([P24openname(1:8) '(JD' num2str(jd) ')' ' Lower Tiltmeter'])
%     xlabel('Day Time (hour)');   
%     axes(P23bottomhaxes(1))
%     xlim([6 20])      
%     set(gca,'XTick',6:1:20)
%     ylabel('P23 Tilt (\mu rad)')
%     ylim([-tiltboundaryb 30])      
%     set(gca,'YTick',-tiltboundaryb:tiltinterval:30)
%     axes(P23bottomhaxes(2))
%     xlim([6 20])      
%     set(gca,'XTick',6:1:20)
%     ylabel('Wind Speed (m/s)')
%     ylim([-windboundaryb 5])          
%     set(gca,'YTick',-windboundaryb:windinterval:5) 
%     set(P23bwindplot,'LineStyle',':','LineWidth',2)
%     Grid on   
%     
% %     subplot(2,1,2)
% %     [P24bottomhaxes, P24bottomtiltplot,P24bwindplot] = plotyy(P24tiltjd,P24bottomtilt, windjd,-windsd);        
% %     text(P24bmaxpoint,max(P24bottomtilt),['\leftarrow ' num2str(max(P24bottomtilt))],...
% %         'HorizontalAlignment','left')    
% %     text(P24bminpoint,min(P24bottomtilt),[num2str(min(P24bottomtilt)) '\rightarrow'],...
% %         'HorizontalAlignment','right')        
% %     xlabel('Day Time (hour)');      
% %     axes(P24bottomhaxes(1))
% %     xlim([0 24])      
% %     set(gca,'XTick',0:1:24)
% %     ylabel('P24 Tilt (\mu rad)')
% %     ylim([-tiltboundaryb 50])      
% %     set(gca,'YTick',-tiltboundaryb:tiltinterval:50)
% %     axes(P24bottomhaxes(2))
% %     xlim([0 24])      
% %     set(gca,'XTick',0:1:24)
% %     ylabel('Wind Speed (m/s)')
% %     ylim([-windboundaryb 5])          
% %     set(gca,'YTick',-windboundaryb:windinterval:5) 
% %     set(P24bwindplot,'LineStyle',':','LineWidth',2)
% %     Grid on 
% %     
% %     savefile = ['C:\Users\lian\Desktop\2\' P24openname(1:11) 'Bottomtilt'];
% %     saveas(P23bottomtiltplot, savefile , 'jpg')     
% %     %close all
%          
% end    


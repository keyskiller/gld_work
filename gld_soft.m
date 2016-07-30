function varargout = gld_soft(varargin)
% GLD_SOFT MATLAB code for gld_soft.fig
%      GLD_SOFT, by itself, creates a new GLD_SOFT or raises the existing
%      singleton*.
%
%      H = GLD_SOFT returns the handle to a new GLD_SOFT or the handle to
%      the existing singleton*.
%
%      GLD_SOFT('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in GLD_SOFT.M with the given input arguments.
%
%      GLD_SOFT('Property','Value',...) creates a new GLD_SOFT or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before gld_soft_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to gld_soft_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help gld_soft

% Last Modified by GUIDE v2.5 28-Jul-2016 22:36:15

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @gld_soft_OpeningFcn, ...
                   'gui_OutputFcn',  @gld_soft_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before gld_soft is made visible.
function gld_soft_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to gld_soft (see VARARGIN)

% Choose default command line output for gld_soft
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes gld_soft wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = gld_soft_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
%% 先获取各项数据
%当事人信息提取，公民部分
xingmin = (get(handles.edit1,'string'));%当事人姓名
xingbie = (get(handles.edit2,'string'));%性别
shenfenzhenghao = (get(handles.edit3,'string'));%身份证号
zhuzhi = (get(handles.edit4,'string'));%住址
zhiye = (get(handles.edit5,'string'));%职业
nianling = (get(handles.edit30,'string'));%年龄
gongmindianhua = (get(handles.edit31,'string'));%当事人电话
anjianguanxi = (get(handles.edit39,'string'));%案件关系
danweizhiwu = (get(handles.edit40,'string'));%单位职务

%当事人信息提取，法人部分
mingcheng = (get(handles.edit6,'string'));%法人名称
fadingdaibiaoren = (get(handles.edit7,'string'));%法人法定代表人
dizhi = (get(handles.edit8,'string'));%地址
farendianhua = (get(handles.edit32,'string'));%法人电话


% 违法时间提取
weifanian = (get(handles.edit9,'string'));
weifayue = (get(handles.edit10,'string'));
weifari = (get(handles.edit11,'string'));
weifashi = (get(handles.edit12,'string'));
weifafen = (get(handles.edit13,'string'));

%违法路线
shendao = (get(handles.edit49,'string'));%省道
chuzi = (get(handles.edit50,'string'));%处自
fangxiangxiang = (get(handles.edit51,'string'));%方向向
fangxiangxingshi = (get(handles.edit52,'string'));%方向行驶

%车辆信息提取
chepaihao = (get(handles.edit14,'string'));%车牌号
zhoushuliang = (get(handles.edit15,'string'));%轴数量
huocheleixing = (get(handles.edit16,'string'));%货车类型
huowuleixing = (get(handles.edit17,'string'));%货物类型
chehuozongliang = (get(handles.edit18,'string'));%车货总量
chaoxiandunshu = (get(handles.edit19,'string'));%超限吨数
jiashiyuanzhengjianleixing  = (get(handles.edit20,'string'));%驾驶员证件类型
cheliangzhengjianleixing  = (get(handles.edit21,'string'));%车辆证件类型
xingzhengchufaleixing  = (get(handles.edit22,'string'));%行政处罚类型
yinhangzhanghao  = (get(handles.edit23,'string'));%银行账号

% 处罚时间提取
chufanian = (get(handles.edit24,'string'));
chufayue = (get(handles.edit25,'string'));
chufari = (get(handles.edit26,'string'));
chufashi = (get(handles.edit27,'string'));
chufafen = (get(handles.edit28,'string'));

% 执法人员信息提取
zhifadidian = (get(handles.edit33,'string'));%执法地点
zhifarenyuan1 = (get(handles.edit34,'string'));%执法人员1
zhifarenyuan2 = (get(handles.edit35,'string'));%执法人员2
jiluren = (get(handles.edit36,'string'));%执法人员2
zhifazhenghao1 = (get(handles.edit37,'string'));%执法证号1
zhifazhenghao2 = (get(handles.edit38,'string'));%执法证号2
zhifanian = (get(handles.edit42,'string'));%执法时间
zhifayue = (get(handles.edit43,'string'));
zhifari = (get(handles.edit44,'string'));
zhifashi1 = (get(handles.edit45,'string'));
zhifafen1 = (get(handles.edit46,'string'));
zhifashi2 = (get(handles.edit47,'string'));
zhifafen2 = (get(handles.edit48,'string'));




%用Matlab生成Word文档
%用Matlab编了一段程序，可以生成Word文档，文档中含有表格，代码如下：
filespec = ['自动测试报告' datestr(now,30) '.doc'];
try
    Word=actxGetRunningServer('Word.Application');
catch
    Word = actxserver('Word.Application'); 
end;
set(Word, 'Visible', 1);
documents = Word.Documents;
if exist(filespec,'file')
    document = invoke(documents,'Open',filespec);    
else
    document = invoke(documents, 'Add');
    document.SaveAs2(filespec);
end
content = document.Content;
duplicate = content.Duplicate;
inlineshapes = content.InlineShapes;
selection = Word.Selection;
paragraphformat = selection.ParagraphFormat;
%页面设置
document.PageSetup.TopMargin = 60;
document.PageSetup.BottomMargin = 45;
document.PageSetup.LeftMargin = 40;
document.PageSetup.RightMargin = 40;



% 第一页
set(content, 'Start',0);
title='行政处罚决定书';
set(content, 'Text',title);
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
rr=document.Range(0,7);%选择文本
rr.Font.Size=16;%设置文本字体大小
rr.Font.Bold=4;%设置文本粗体
rr.Font.Name = '宋体';%选择字体
% rr.Font.Bold = 'wdToggle';
end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
selection.Font.Size=10;
selection.MoveDown;
selection.TypeParagraph;
 set(paragraphformat, 'Alignment','wdAlignParagraphRight');
 
 text1='鄂当路政  罚（    ）     号';
set(selection, 'Text',text1);
dd=document.Range(8,14);%选择文本
 dd.Font.underline = 2;
 dd.Font.Size=12;%设置文本字体大小
dd.Font.Bold=4;%设置文本粗体

dd=document.Range(15,28);%选择文本
%  dd.Font.underline = 2;
 dd.Font.Size=15;%设置文本字体大小
% dd.Font.Bold=4;%设置文本粗体

% 这段看起来时回车的意思
end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 selection.TypeParagraph;
 
selection.Font.Size=15;
Tables=document.Tables.Add(selection.Range,4,8);
 
%设置边框
DTI=document.Tables.Item(1);
DTI.Borders.OutsideLineStyle='wdLineStyleSingle';
% DTI.Borders.OutsideLineWidth='wdLineWidth150pt';
DTI.Borders.InsideLineStyle='wdLineStyleSingle';
% DTI.Borders.InsideLineWidth='wdLineWidth150pt';
DTI.Rows.Alignment='wdAlignRowCenter';
% end_of_doc = get(content,'end');
% set(selection,'Start',end_of_doc);
% selection.TypeParagraph;
% set(selection, 'Text','主管签字：            年    月    日');
% set(paragraphformat, 'Alignment','wdAlignParagraphRight');
% end_of_doc = get(content,'end');
% set(selection,'Start',end_of_doc);
%  DTI.Rows.Item(5).Borders.Item(1).LineStyle='wdLineStyleNone';
column_width=[30, 70, 50, 100, 50, 45, 45, 130];
row_height=[41, 41, 41, 41];
for i = 1 : 8
    DTI.Columns.Item(i).Width =column_width(i);
end
for i=1:4
    DTI.Rows.Item(i).Height =row_height(i);
end
for i = 1 : 4
    for j = 1 : 8
        DTI.Cell(i, j).Range.ParagraphFormat.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i, j).VerticalAlignment='wdCellAlignVerticalCenter';
    end
end
% DTI.Cell(4,1).Range.ParagraphFormat.Alignment='wdAlignParagraphLeft';
 DTI.Cell(1, 1).Merge(DTI.Cell(4, 1));
 DTI.Cell(1, 2).Merge(DTI.Cell(2, 2));
 DTI.Cell(3, 2).Merge(DTI.Cell(4, 2));
 DTI.Cell(2, 4).Merge(DTI.Cell(2, 6));
 DTI.Cell(3, 4).Merge(DTI.Cell(3, 5));
  DTI.Cell(4, 4).Merge(DTI.Cell(4, 8));
  DTI.Cell(3, 5).Merge(DTI.Cell(3, 6));
  
  DTI.Cell(1, 1).Range.Text = '当事人';
  DTI.Cell(1, 2).Range.Text = '公民';
  DTI.Cell(1, 3).Range.Text = '姓名';
  DTI.Cell(1, 4).Range.Text = xingmin;
    DTI.Cell(1, 5).Range.Text = '性别';
   DTI.Cell(1, 6).Range.Text = xingbie;
  DTI.Cell(1, 7).Range.Text = '身份证号';
  DTI.Cell(1, 8).Range.Text = shenfenzhenghao;
   DTI.Cell(1, 8).Range.Font.Size = 12 ;%给该单元格改字体大小
  DTI.Cell(2, 3).Range.Text = '住址';
  DTI.Cell(2, 4).Range.Text = zhuzhi;
  DTI.Cell(2, 4).Range.Font.Size = 12 ;
  DTI.Cell(2, 5).Range.Text = '职业';
  DTI.Cell(2, 6).Range.Text = zhiye;
  DTI.Cell(2, 6).Range.Font.Size = 12 ;
  DTI.Cell(3, 2).Range.Text = '法人或其他组织';
  DTI.Cell(3, 3).Range.Text = '名称';
  DTI.Cell(3, 4).Range.Text = mingcheng;
  DTI.Cell(3, 4).Range.Font.Size = 12 ;
  DTI.Cell(4, 3).Range.Text = '地址';
  DTI.Cell(4, 4).Range.Text = dizhi;
  DTI.Cell(4, 4).Range.Font.Size = 12 ;
  DTI.Cell(3, 5).Range.Text = '法定代表人';
  DTI.Cell(3, 6).Range.Text = fadingdaibiaoren;
  
  % 这段看起来时回车的意思
end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
 selection.TypeParagraph;
  set(paragraphformat, 'LineSpacingRule','wdLineSpaceExactly');
 set(paragraphformat, 'LineSpacing',23);


text2='    违法事实及证据：';%
set(selection, 'Text',text2);
selection.Font.Size=15;
end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 
text3=['当事人'    xingmin      '于' weifanian '年' weifayue '月' weifari  '日'  weifashi  '时' weifafen   ...
      '分，驾驶牌号为' chepaihao '的' zhoushuliang  ' 轴 ' huocheleixing  '运输'  huowuleixing ...
      '在 ' shendao  '省道上行驶，经当阳市公路管理局路政人员检测，车货总重' chehuozongliang '吨，超限 ' chaoxiandunshu ...
      '吨，属擅自超限运输行驶公路，前述事实由《现场笔录》、《询问笔录》、《检测磅单》、《 驾驶员' jiashiyuanzhengjianleixing ...
      '  复印件》、《车辆  '    cheliangzhengjianleixing   ' 复印件》予以佐证'  xingmin  ' 对违法事实予以承认。'];%
set(selection, 'Text',text3);
selection.Font.Size=13;
selection.Font.Bold=4;%设置文本粗体
set(paragraphformat, 'Alignment','wdAlignParagraphJustify');
selection.Font.underline = 2;
  set(paragraphformat, 'LineSpacingRule','wdLineSpaceExactly');
 set(paragraphformat, 'LineSpacing',23);
 
end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 selection.TypeParagraph;
 text4=['    以上事实违反了《中华人民共和国公路法》第四十九条 、第五十条，《公路安全保护条例》第三十三条第一款 的规定，依据 《中华人民共和国公路法》第七十六条第五项、《公路安全保护条例》第六十四条规定，决定给予 ' xingzhengchufaleixing ' 的行政处罚。'];%
set(selection, 'Text',text4);
selection.Font.Size=13;
selection.Font.Bold=4;%设置文本粗体
  set(paragraphformat, 'LineSpacingRule','wdLineSpaceExactly');
 set(paragraphformat, 'LineSpacing',23);


end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 selection.TypeParagraph;
 text5=['    处以罚款的，罚款自收到本决定书之日起15日内缴至 当阳市建设银行营业部   ，账号为：' yinhangzhanghao '到期不缴的依法每日按罚款数额的3%加处罚款。'];%
set(selection, 'Text',text5);
selection.Font.Size=13;
selection.Font.Bold=4;%设置文本粗体

end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 selection.TypeParagraph;
 text6=['    如果不服本处罚决定，可以依法在60日内向   当阳市交通运输局  申请行政复议，或者在三个月内依法向人民法院提起行政诉讼，但本决定不停止执行，法律另有规定的除外。逾期不申请行政复议、不提起行政诉讼又不履行的，本机关将依法申请人民法院强制执行或者依照有关规定强制执行。'];%
set(selection, 'Text',text6);
selection.Font.Size=13;
selection.Font.Bold=4;%设置文本粗体

end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 selection.TypeParagraph;
 selection.TypeParagraph;
text7=[ chufanian '   年' chufayue '月'  chufari  '     日' ];%
set(selection, 'Text',text7);
selection.Font.Size=14;
selection.Font.Name = '宋体';%选择字体
 selection.Font.Bold=0;%设置文本正常
set(paragraphformat, 'Alignment','wdAlignParagraphRight');

end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 selection.TypeParagraph;
 selection.TypeParagraph;
text8=[ '（本文书一式两份：一份存根，一份交当事人或其代理人。）' ];%
set(selection, 'Text',text8);
selection.Font.Size=14;
selection.Font.Name = '宋体';%选择字体
 selection.Font.Bold=0;%设置文本正常
set(paragraphformat, 'Alignment','wdAlignParagraphLeft');

%第二页
selection.MoveDown;
selection.TypeParagraph;
selection.TypeParagraph;
selection.TypeParagraph;
 text21='立  案  审  批  表';
set(selection, 'Text',text21);


 selection.Font.Size=16;%设置文本字体大小
selection.Font.Bold=4;%设置文本粗体
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');

selection.MoveDown;
selection.TypeParagraph;
 set(paragraphformat, 'Alignment','wdAlignParagraphRight');
 
 text_temp='鄂当路政';%  罚（    ）     号';
set(selection, 'Text',text_temp);

 selection.Font.underline = 2;
 selection.Font.Size=12;%设置文本字体大小
 selection.Font.Bold=4;%设置文本粗体

end_of_doc = get(content,'end');%光标位置移动到最后
 set(selection,'Start',end_of_doc);

 text_temp='  罚（    ）     号';
 set(selection, 'Text',text_temp);
%新的文字的格式需要重新设置，否则就是沿用之前的格式
 selection.Font.Bold=0;%设置文本粗体
 selection.Font.underline = 0;
 selection.Font.Size=15;%设置文本字体大小
 end_of_doc = get(content,'end');%光标位置移动到最后,每个内容输入完毕后都应该作这个操作
 set(selection,'Start',end_of_doc);



%开始画第二页的图
 selection.Font.Bold=0;%设置文本粗体
 selection.Font.underline = 0;
 selection.Font.Size=15;%设置文本字体大小
 Tables=document.Tables.Add(selection.Range,10,8);
 
%设置边框
DTI=document.Tables.Item(2);
DTI.Borders.OutsideLineStyle='wdLineStyleSingle';
% DTI.Borders.OutsideLineWidth='wdLineWidth150pt';
DTI.Borders.InsideLineStyle='wdLineStyleSingle';
% DTI.Borders.InsideLineWidth='wdLineWidth150pt';
DTI.Rows.Alignment='wdAlignRowCenter';

column_width=[50, 50, 50, 100, 50, 70, 60, 100];
row_height=[60, 35, 31, 71, 51, 51, 100, 100, 100, 30 ];
for i = 1 : 8
    DTI.Columns.Item(i).Width =column_width(i);
end
for i=1:10
    DTI.Rows.Item(i).Height =row_height(i);
end
for i = 1 : 10
    for j = 1 : 8
        DTI.Cell(i, j).Range.ParagraphFormat.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i, j).VerticalAlignment='wdCellAlignVerticalCenter';
    end
end
 DTI.Cell(1, 2).Merge(DTI.Cell(1, 5));
 DTI.Cell(1, 4).Merge(DTI.Cell(1, 5));
 DTI.Cell(2, 2).Merge(DTI.Cell(2, 8)); 
 DTI.Cell(3, 1).Merge(DTI.Cell(6, 1)); 
 DTI.Cell(3, 2).Merge(DTI.Cell(4, 2)); 
 DTI.Cell(5, 2).Merge(DTI.Cell(6, 2)); 
 DTI.Cell(5, 4).Merge(DTI.Cell(5, 6)); 
 DTI.Cell(6, 4).Merge(DTI.Cell(6, 6)); 
 DTI.Cell(7, 2).Merge(DTI.Cell(7, 8)); 
 DTI.Cell(8, 2).Merge(DTI.Cell(8, 5)); 
 DTI.Cell(8, 4).Merge(DTI.Cell(8, 5)); 
 DTI.Cell(9, 2).Merge(DTI.Cell(9, 8));
 DTI.Cell(10, 2).Merge(DTI.Cell(10, 8));
 
  DTI.Cell(1, 1).Range.Text = '案件来源';
  DTI.Cell(1, 2).Range.Text = '执法检查发现';
  DTI.Cell(1, 2).Range.Font.Size = 13 ;%给该单元格改字体大小
  DTI.Cell(1, 2).Range.Font.Bold = 2 ;%给该单元格改字体大小
  DTI.Cell(1, 3).Range.Text = '受案时间';
  DTI.Cell(1, 4).Range.Text = [weifanian '年' weifayue '月' weifari '日' weifashi '时' weifafen '分' ];
  DTI.Cell(2, 1).Range.Text = '案由';
  DTI.Cell(2, 2).Range.Text = '涉嫌擅自超限运输行驶公路案';
  DTI.Cell(2, 2).Range.Font.Size = 13 ;%给该单元格改字体大小
  DTI.Cell(2, 2).Range.Font.Bold = 2 ;%给该单元格改字体大小
  DTI.Cell(3, 1).Range.Text = '当事人基本情况';
  DTI.Cell(3, 2).Range.Text = '公民';
  DTI.Cell(3, 3).Range.Text = '姓名';
  DTI.Cell(3, 4).Range.Text = xingmin;
  DTI.Cell(3, 5).Range.Text = '性别';
  DTI.Cell(3, 6).Range.Text = xingbie;
  DTI.Cell(3, 7).Range.Text = '年龄';
  DTI.Cell(3, 8).Range.Text = nianling;
  DTI.Cell(4, 3).Range.Text = '住址';
  DTI.Cell(4, 4).Range.Text = zhuzhi;
  DTI.Cell(4, 4).Range.Font.Size = 12 ;%给该单元格改字体大小
  DTI.Cell(4, 5).Range.Text = '身份证号';
  DTI.Cell(4, 6).Range.Text = shenfenzhenghao;
  DTI.Cell(4, 6).Range.Font.Size = 12 ;%给该单元格改字体大小
  DTI.Cell(4, 7).Range.Text = '联系电话';
  DTI.Cell(4, 8).Range.Text = gongmindianhua;
  DTI.Cell(5, 2).Range.Text = '法人或其他组织';
  DTI.Cell(5, 3).Range.Text = '名称';
  DTI.Cell(5, 4).Range.Text = mingcheng;
  DTI.Cell(5, 5).Range.Text = '法定代表人';
  DTI.Cell(5, 6).Range.Text = fadingdaibiaoren;
  DTI.Cell(6, 3).Range.Text = '地址';
  DTI.Cell(6, 4).Range.Text = dizhi;
  DTI.Cell(6, 4).Range.Font.Size = 12 ;%给该单元格改字体大小
  DTI.Cell(6, 5).Range.Text = '联系电话';
  DTI.Cell(6, 6).Range.Text = farendianhua;
  DTI.Cell(7, 1).Range.Text = '案件基本情况';
  DTI.Cell(7, 2).Range.Text = [weifanian '年' weifayue '月' weifari '日' weifashi '时' weifafen '分' ...
                                '，当阳市公路管理局路政执法人员' zhifarenyuan1  '、 ' zhifarenyuan2 ...
                                '  在例行检测中发现, 驾驶员' xingmin '驾驶车牌号为 ' chepaihao      ...
                                ' 的 ' zhoushuliang   ' 轴' huocheleixing    '货车运输 ' huowuleixing    ...
                                ' ,在' shendao  '省道' chuzi  '处自' fangxiangxiang  '方向向 ' fangxiangxingshi   ...
                                '   方向行驶。经依法检测，该车车货总重（长、宽、高） '  chehuozongliang   ...
                                '  吨（米），涉嫌擅自超限运输行驶公路。'];
  DTI.Cell(7, 2).Range.Font.Size = 12 ;%给该单元格改字体大小
  DTI.Cell(8, 1).Range.Text = '立案依据';
  DTI.Cell(8, 2).Range.Text = '《中华人民共和国公路法》第四十九条 、第五十条，《公路安全保护条例》第三十三条第一款。';
  DTI.Cell(8, 2).Range.Font.Size = 13 ;%给该单元格改字体大小
  DTI.Cell(8, 2).Range.Font.Bold = 2 ;%给该单元格改字体大小
  DTI.Cell(8, 3).Range.Text = '受案机构意见';
  DTI.Cell(8, 4).Range.Text = '                                    签名：               时间：    年 月 日';
  DTI.Cell(8, 4).Range.ParagraphFormat.Alignment='wdAlignParagraphLeft';
  DTI.Cell(8, 4).Range.Font.Size = 13 ;%给该单元格改字体大小
  DTI.Cell(8, 4).VerticalAlignment='wdCellAlignVerticalBottom';
  DTI.Cell(9, 1).Range.Text = '负责审批人意见';
  DTI.Cell(9, 2).Range.Text = '                                      签名：                  .        时间：      年   月    日';
  DTI.Cell(9, 2).Range.ParagraphFormat.Alignment='wdAlignParagraphRight';
  DTI.Cell(9, 2).VerticalAlignment='wdCellAlignVerticalBottom';
  DTI.Cell(9, 2).Range.Font.Size = 13 ;%给该单元格改字体大小
  DTI.Cell(10, 1).Range.Text = '备注';
  DTI.Cell(10, 2).Range.Text = '此项无内容';
  DTI.Cell(10, 2).Range.Font.Size = 13 ;%给该单元格改字体大小
  DTI.Cell(10, 2).Range.Font.Bold = 2 ;%给该单元格改字体大小
  % 加入需要修改格式和填空的内容
  
  
  %第三页
   end_of_doc = get(content,'end');%光标位置移动到最后,每个内容输入完毕后都应该作这个操作
 set(selection,'Start',end_of_doc);
  selection.MoveDown;
selection.TypeParagraph;
selection.TypeParagraph;
selection.TypeParagraph;
 text_temp='现  场  笔  录';
set(selection, 'Text',text_temp);


 selection.Font.Size=16;%设置文本字体大小
selection.Font.Bold=4;%设置文本粗体
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');

selection.MoveDown;
selection.TypeParagraph;
 
 end_of_doc = get(content,'end');%光标位置移动到最后,每个内容输入完毕后都应该作这个操作
 set(selection,'Start',end_of_doc);
  
  %开始画第三页的图
 selection.Font.Bold=0;%设置文本粗体
 selection.Font.underline = 0;
 selection.Font.Size=15;%设置文本字体大小
 Tables=document.Tables.Add(selection.Range,11,7);
 
%设置边框
DTI=document.Tables.Item(3);
DTI.Borders.OutsideLineStyle='wdLineStyleSingle';
% DTI.Borders.OutsideLineWidth='wdLineWidth150pt';
DTI.Borders.InsideLineStyle='wdLineStyleSingle';
% DTI.Borders.InsideLineWidth='wdLineWidth150pt';
DTI.Rows.Alignment='wdAlignRowCenter';

column_width=[50, 80, 80, 40, 80, 60, 120];
row_height=[40, 30, 30, 41, 41, 41, 41, 41, 250, 40 , 40 ];
for i = 1 : 7
    DTI.Columns.Item(i).Width =column_width(i);
end
for i=1:11
    DTI.Rows.Item(i).Height =row_height(i);
end
for i = 1 : 11
    for j = 1 : 7
        DTI.Cell(i, j).Range.ParagraphFormat.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i, j).VerticalAlignment='wdCellAlignVerticalCenter';
    end
end
  DTI.Cell(1, 2).Merge(DTI.Cell(1, 3));
  DTI.Cell(1, 3).Merge(DTI.Cell(1, 4));
  DTI.Cell(1, 4).Merge(DTI.Cell(1, 5));
  DTI.Cell(2, 1).Merge(DTI.Cell(3, 1));
  DTI.Cell(2, 3).Merge(DTI.Cell(3, 3));
  DTI.Cell(2, 4).Merge(DTI.Cell(2, 5));
  DTI.Cell(3, 4).Merge(DTI.Cell(3, 5));
  DTI.Cell(2, 5).Merge(DTI.Cell(3, 5));
  DTI.Cell(2, 6).Merge(DTI.Cell(3, 6));
  DTI.Cell(4, 1).Merge(DTI.Cell(8, 1));
  DTI.Cell(9, 1).Merge(DTI.Cell(11, 1));
  DTI.Cell(4, 3).Merge(DTI.Cell(4, 4));
  DTI.Cell(4, 5).Merge(DTI.Cell(4, 6));
  DTI.Cell(5, 3).Merge(DTI.Cell(5, 4));
  DTI.Cell(5, 5).Merge(DTI.Cell(5, 6));
  DTI.Cell(6, 3).Merge(DTI.Cell(6, 4));
  DTI.Cell(6, 5).Merge(DTI.Cell(6, 6));
  DTI.Cell(7, 3).Merge(DTI.Cell(7, 7));
  DTI.Cell(8, 3).Merge(DTI.Cell(8, 4));
  DTI.Cell(8, 5).Merge(DTI.Cell(8, 6));
  DTI.Cell(9, 2).Merge(DTI.Cell(9, 7));
  DTI.Cell(10, 2).Merge(DTI.Cell(10, 7));
  DTI.Cell(11, 2).Merge(DTI.Cell(11, 7));
  
  DTI.Cell(1, 1).Range.Text = '执法地点';
  DTI.Cell(1, 2).Range.Text = zhifadidian ;
  DTI.Cell(1, 3).Range.Text = '执法时间';
  DTI.Cell(1, 4).Range.Text = [zhifanian '年' zhifayue  '月' zhifari '日'...
                               zhifashi1 '时'  zhifafen1 '分至' zhifashi2 '时' zhifafen2 '分' ];
  DTI.Cell(2, 1).Range.Text = '执法人员';
  DTI.Cell(2, 2).Range.Text = zhifarenyuan1 ;
  DTI.Cell(3, 2).Range.Text = zhifarenyuan2 ;
  DTI.Cell(2, 3).Range.Text = '执法证号';
  DTI.Cell(2, 4).Range.Text = zhifazhenghao1 ;
  DTI.Cell(3, 4).Range.Text = zhifazhenghao2 ;
  DTI.Cell(2, 5).Range.Text = '记录人';
  DTI.Cell(2, 6).Range.Text = jiluren ;
  DTI.Cell(4, 1).Range.Text = '现场人员基本情况';
  DTI.Cell(4, 2).Range.Text = '姓   名';
  DTI.Cell(4, 3).Range.Text =  xingmin ;
  DTI.Cell(4, 4).Range.Text = '性   别';
  DTI.Cell(4, 5).Range.Text =  xingbie;
  DTI.Cell(5, 2).Range.Text = '身份证号';
  DTI.Cell(5, 3).Range.Text = shenfenzhenghao ;
  DTI.Cell(5, 4).Range.Text = '与案件关系';
  DTI.Cell(5, 5).Range.Text = anjianguanxi ;
  DTI.Cell(6, 2).Range.Text = '单位及职务';
  DTI.Cell(6, 3).Range.Text = danweizhiwu ;
  DTI.Cell(6, 4).Range.Text = '联系电话';
  DTI.Cell(6, 5).Range.Text = gongmindianhua ;
  DTI.Cell(7, 2).Range.Text = '联系地址';
  DTI.Cell(7, 3).Range.Text = zhuzhi ;
  DTI.Cell(8, 2).Range.Text = '车（船）号';
  DTI.Cell(8, 3).Range.Text = chepaihao ;
  DTI.Cell(8, 4).Range.Text = '车（船）型';
  DTI.Cell(8, 5).Range.Text = huocheleixing ;
  DTI.Cell(9, 1).Range.Text = '主要内容';
  DTI.Cell(9, 2).Range.Text = ['在检查中发现：驾驶员'  xingmin  '于  '...
                                weifanian  ' 年 '  weifayue  ' 月 '  weifari ...
                                ' 日 ' weifashi   ' 时 ' weifafen  ...
                                ' 分，驾驶 '     ' 牌号为 '  chepaihao ...
                                '的 ' zhoushuliang  ' 轴 ' huocheleixing '  货车运输 ' huowuleixing ...
                                ' 在 ' shendao  ' 省道 ' chuzi  ' 处自 ' fangxiangxiang  '方向向 ' fangxiangxingshi  ' 方向行驶，当阳市公路管理局路政人员依法对其进行了检测，经检测，车货总重 '...
                                chehuozongliang  '  吨。                                        （以下无内容）    '...
                                '                                 __________________________________________________________________________ 上述笔录我已看过（或已向我宣读过），情况属实无误。　'...
                                 '现场人员签名：             时间：   年  月  日'];
  DTI.Cell(10, 2).Range.Text = '备注';
  DTI.Cell(10, 2).Range.ParagraphFormat.Alignment='wdAlignParagraphLeft';
  DTI.Cell(11, 2).Range.Text = '执法人员签名：                         时间：   年  月  日';
  DTI.Cell(11, 2).Range.ParagraphFormat.Alignment='wdAlignParagraphLeft';
  end_of_doc = get(content,'end');%光标位置移动到最后,每个内容输入完毕后都应该作这个操作
  set(selection,'Start',end_of_doc);
%   set(selection, 'Text',DTI.Cell(11, 2).Range.Text);
  selection.MoveDown;
  selection.TypeParagraph;
  
  %第四页开始
   text_temp='询   问   笔   录';
   set(selection, 'Text',text_temp);
   selection.Font.Size=16;%设置文本字体大小
   selection.Font.Bold=4;%设置文本粗体
   set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
   
    end_of_doc = get(content,'end');%光标位置移动到最后,每个内容输入完毕后都应该作这个操作
    set(selection,'Start',end_of_doc);
    selection.MoveDown;
    selection.TypeParagraph;
  
    %开始画第四页的图
 selection.Font.Bold=0;%设置文本粗体
 selection.Font.underline = 0;
 selection.Font.Size=15;%设置文本字体大小
 Tables=document.Tables.Add(selection.Range,8,4);
 
%设置边框
DTI=document.Tables.Item(4);
DTI.Borders.OutsideLineStyle='wdLineStyleSingle';
% DTI.Borders.OutsideLineWidth='wdLineWidth150pt';
DTI.Borders.InsideLineStyle='wdLineStyleSingle';
% DTI.Borders.InsideLineWidth='wdLineWidth150pt';
DTI.Rows.Alignment='wdAlignRowCenter';

column_width=[120, 180, 100, 140];
row_height=[25, 25, 25, 25, 25, 25, 25, 25, 25 ];
for i = 1 : 4
    DTI.Columns.Item(i).Width =column_width(i);
end
for i=1:8
    DTI.Rows.Item(i).Height =row_height(i);
end
for i = 1 : 8
    for j = 1 : 4
        DTI.Cell(i, j).Range.ParagraphFormat.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i, j).VerticalAlignment='wdCellAlignVerticalCenter';
    end
end
  
DTI.Cell(1, 2).Merge(DTI.Cell(1, 4));
DTI.Cell(2, 2).Merge(DTI.Cell(2, 4));
DTI.Cell(7, 2).Merge(DTI.Cell(7, 4));
DTI.Cell(8, 2).Merge(DTI.Cell(8, 4));
  
DTI.Cell(1, 1).Range.Text = '时间';
DTI.Cell(2, 1).Range.Text = '询问地点';
DTI.Cell(3, 1).Range.Text = '询问人';
DTI.Cell(3, 3).Range.Text = '记录人';
DTI.Cell(4, 1).Range.Text = '被询问人';
DTI.Cell(4, 2).Range.Text = xingmin ;
DTI.Cell(4, 3).Range.Text = '与案件关系';
DTI.Cell(4, 4).Range.Text = anjianguanxi ;
DTI.Cell(5, 1).Range.Text = '性别';
DTI.Cell(5, 2).Range.Text = xingbie ;
DTI.Cell(5, 3).Range.Text = '年龄';
DTI.Cell(5, 4).Range.Text = nianling ;
DTI.Cell(6, 1).Range.Text = '身份证号';
DTI.Cell(6, 2).Range.Text = shenfenzhenghao ;
DTI.Cell(6, 3).Range.Text = '电话';
DTI.Cell(6, 4).Range.Text = gongmindianhua ;
DTI.Cell(7, 1).Range.Text = '工作单位及职务';
DTI.Cell(7, 2).Range.Text = danweizhiwu ;
DTI.Cell(8, 1).Range.Text = '联系地址';
DTI.Cell(8, 2).Range.Text = zhuzhi ;
  
  end_of_doc = get(content,'end');%光标位置移动到最后,每个内容输入完毕后都应该作这个操作
  set(selection,'Start',end_of_doc);
   set(paragraphformat, 'LineSpacing',20);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['    我们是当阳市公路管理局路政执法人员' zhifarenyuan1  ' 、' zhifarenyuan2  ' ，'...
        '这是我们的证件，执法证件号码分别是 ' zhifazhenghao1  '  、 ' zhifazhenghao2  '  ，'...
        '请你确认。现依法向你询问，请如实回答所问问题，执法人员与你有直接利害关系的，你可以申请回避。'];%
  set(selection, 'Text',text2);
  selection.Font.Size=15;
  
  %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：你是否申请执法人员回避？                                         ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答：否                                        '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体
  
    %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：你叫什么名字?                                          ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答：' xingmin ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体
  
      %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：现年多少岁?                                            ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答： '  nianling  ' 岁 '   ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体
  
        %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：你的现在住在什么地方？                                            ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答：' zhuzhi  ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体
  
 %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：你能出示一下身份证件吗？                                            ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答：身份证号是' shenfenzhenghao   ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体 
  
   %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：你在什么单位工作？                                             ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答： ' danweizhiwu  ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体 
  
     %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['问：' weifanian  ' 年  ' weifayue  ' 月 ' weifari   ' 日车牌号为 ' chepaihao  ' 货车在' shendao  '省道上进行货物运输时是你驾驶的吗？  '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='答：是的';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体 
  
         %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：车主是谁？  ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答： ' ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体 
  
      end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
   set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='被询问人签名及时间：                         询问人签名及时间';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%设置文本粗体
  
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='____________________                       ______________________';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%设置文本粗体
  
  

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='备注：此项无内容';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%设置文本粗体

  
    %第五页开始
  
    
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='续页';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%设置文本粗体
  

  
  %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：车号是多少？ ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答： ' chepaihao  ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体 
  
  %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：今天运输的是什么货物？ ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答： ' huowuleixing  ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体 
  
  %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：装了多少吨？ ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答：否   '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体
  
  %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：从什么地方拉来的？ ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答：否                                        '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体
  
  %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：目的地在什么地方？ ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答：否                                        '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体
  
  %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：这趟运费是多少？ ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答：否                                        '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体
  
  %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：你是从哪条路行驶到这里的？  ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答：否                                        '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体
  
  %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['问：经我们称重仪检测，车货总重 ' chehuozongliang  '吨，对此检测结果有无异议？  '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答： ''没有异议'];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体
  
  %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：其他还有补充的吗？  ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答：''没有'];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体
  
  %问答类型模板
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='问：以上记录是否属实，请你再看一下，如属实，请签字。 ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%设置文本粗体

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['答：_____________________________________________________'...
      '____________________ _______________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%设置文本粗体
  
        end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
   set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='被询问人签名及时间：                         询问人签名及时间';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%设置文本粗体

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='____________________                       ______________________';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%设置文本粗体
  
  

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='备注：此项无内容';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%设置文本粗体
  
  %第六页开始
  
    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  
  text_temp='勘验（检查）笔录';
  set(selection, 'Text',text_temp);


  selection.Font.Size=16;%设置文本字体大小
  selection.Font.Bold=4;%设置文本粗体
  set(paragraphformat, 'Alignment','wdAlignParagraphCenter');

  selection.MoveDown;
  selection.TypeParagraph;
  
  text_temp='案由： ';
  set(selection, 'Text',text_temp);
  selection.Font.Size=15;%设置文本字体大小
  selection.Font.Bold=0;%设置文本粗体
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  
  text_temp='_______________________涉嫌擅自超限运输行驶公路案__________________ ';
  set(selection, 'Text',text_temp);
  selection.Font.Size=13;%设置文本字体大小
  selection.Font.Bold=4;%设置文本粗体
  selection.Font.underline = 2;
  
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  
    %开始画第六页的图
  selection.Font.Bold=0;%设置文本粗体
  selection.Font.underline = 0;
  selection.Font.Size=15;%设置文本字体大小
  Tables=document.Tables.Add(selection.Range,9,6);
 
%设置边框
DTI=document.Tables.Item(5);
DTI.Borders.OutsideLineStyle='wdLineStyleSingle';
% DTI.Borders.OutsideLineWidth='wdLineWidth150pt';
DTI.Borders.InsideLineStyle='wdLineStyleSingle';
% DTI.Borders.InsideLineWidth='wdLineWidth150pt';
DTI.Rows.Alignment='wdAlignRowCenter';

column_width=[100,60, 70,100, 100, 80];
row_height=[35, 35, 35, 35, 35, 35, 35, 35, 35 ];
for i = 1 : 6
    DTI.Columns.Item(i).Width =column_width(i);
end
for i=1:9
    DTI.Rows.Item(i).Height =row_height(i);
end
for i = 1 : 9
    for j = 1 : 6
        DTI.Cell(i, j).Range.ParagraphFormat.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i, j).VerticalAlignment='wdCellAlignVerticalCenter';
    end
end
  
DTI.Cell(1, 1).Merge(DTI.Cell(1, 2));
DTI.Cell(1, 2).Merge(DTI.Cell(1, 5));
DTI.Cell(2, 1).Merge(DTI.Cell(2, 2));
DTI.Cell(2, 2).Merge(DTI.Cell(2, 3));
DTI.Cell(6, 2).Merge(DTI.Cell(6, 3));
DTI.Cell(6, 4).Merge(DTI.Cell(6, 5));
DTI.Cell(7, 2).Merge(DTI.Cell(7, 3));
DTI.Cell(7, 4).Merge(DTI.Cell(7, 5));
DTI.Cell(8, 2).Merge(DTI.Cell(8, 3));
DTI.Cell(8, 4).Merge(DTI.Cell(8, 5));
DTI.Cell(9, 2).Merge(DTI.Cell(9, 3));
DTI.Cell(9, 4).Merge(DTI.Cell(9, 5));

DTI.Cell(1, 1).Range.Text = '勘验（检查）时间';
DTI.Cell(2, 1).Range.Text = '勘验（检查）场所';
DTI.Cell(2, 3).Range.Text = '天气情况';
DTI.Cell(3, 1).Range.Text = '勘验（检查）人';
DTI.Cell(3, 3).Range.Text = '单位及职务';
DTI.Cell(3, 5).Range.Text = '执法证号';
DTI.Cell(4, 1).Range.Text = '勘验（检查）人';
DTI.Cell(4, 3).Range.Text = '单位及职务';
DTI.Cell(4, 5).Range.Text = '执法证号';
DTI.Cell(5, 1).Range.Text = '当事人（当事人代理人）姓名';
DTI.Cell(5, 3).Range.Text = '性别';
DTI.Cell(5, 5).Range.Text = '年龄';
DTI.Cell(6, 1).Range.Text = '身份证号';
DTI.Cell(6, 3).Range.Text = '单位及职务';
DTI.Cell(7, 1).Range.Text = '住址';
DTI.Cell(7, 3).Range.Text = '联系电话';
DTI.Cell(8, 1).Range.Text = '被邀请人';
DTI.Cell(8, 3).Range.Text = '单位及职务';
DTI.Cell(9, 1).Range.Text = '记 录 人';
DTI.Cell(9, 3).Range.Text = '单位及职务';


  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  
  text_temp='勘验（检查）情况及结果：';
  set(selection, 'Text',text_temp);
  selection.Font.Size=15;%设置文本字体大小
  selection.Font.Bold=0;%设置文本粗体
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  
  text_temp=['      年   月  日   时   分，驾驶员 '  ' 在场，当阳市公路管理局路政员 '...
            ' 和 ''　在S311省道K234+650处，对          所驾驶的车牌号为 ' ...
            '的货车的车货总重（长、宽、高）依法进行了检测，经检测，该车车货总重       吨。 '...
      '___________________________________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'];
  set(selection, 'Text',text_temp);
  selection.Font.Size=13;%设置文本字体大小
  selection.Font.Bold=4;%设置文本粗体
  selection.Font.underline = 2;
  set(paragraphformat, 'LineSpacing',23);
  
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  text_temp='当事人或其代理人签名：_____________勘验（检查）人签名：_____________';
  set(selection, 'Text',text_temp);
  selection.Font.Size=15;%设置文本字体大小
  selection.Font.Bold=0;%设置文本粗体
  set(paragraphformat, 'LineSpacing',30);
  
  end_of_doc = get(content,'end');%光标移动到最后
  set(selection,'Start',end_of_doc); 
  selection.MoveDown;
  selection.TypeParagraph;
  text_temp='被邀请人签名：_____________________记录人签名：____________________';
  set(selection, 'Text',text_temp);

  
  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  end_of_doc = get(content,'end');%光标移动到最后
  set(selection,'Start',end_of_doc); 
  text_temp='备注：';
  set(selection, 'Text',text_temp);
  
  end_of_doc = get(content,'end');%光标移动到最后
  set(selection,'Start',end_of_doc); 
    selection.MoveDown;
  selection.TypeParagraph;
  
    %第七页开始
  
    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  
  text_temp='现  场  勘  验  图';
  set(selection, 'Text',text_temp);


  selection.Font.Size=16;%设置文本字体大小
  selection.Font.Bold=4;%设置文本粗体
  set(paragraphformat, 'LineSpacing',23);
  set(paragraphformat, 'Alignment','wdAlignParagraphCenter');

  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  
  text_temp='案由： ';
  set(selection, 'Text',text_temp);
  selection.Font.Size=15;%设置文本字体大小
  selection.Font.Bold=0;%设置文本粗体
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  
  text_temp='_______________________涉嫌擅自超限运输行驶公路案__________________ ';
  set(selection, 'Text',text_temp);
  selection.Font.Size=13;%设置文本字体大小
  selection.Font.Bold=4;%设置文本粗体
  selection.Font.underline = 2;
  
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  
    %开始画第七页的图
  selection.Font.Bold=0;%设置文本粗体
  selection.Font.underline = 0;
  selection.Font.Size=15;%设置文本字体大小
  Tables=document.Tables.Add(selection.Range,3,1);
 
%设置边框
DTI=document.Tables.Item(6);
DTI.Borders.OutsideLineStyle='wdLineStyleSingle';
% DTI.Borders.OutsideLineWidth='wdLineWidth150pt';
DTI.Borders.InsideLineStyle='wdLineStyleSingle';
% DTI.Borders.InsideLineWidth='wdLineWidth150pt';
DTI.Rows.Alignment='wdAlignRowCenter';

column_width=510;
row_height=[520, 35, 35 ];

    DTI.Columns.Item(1).Width =column_width;

for i=1:3
    DTI.Rows.Item(i).Height =row_height(i);
end
for i = 1 : 3
    for j = 1 : 1
        DTI.Cell(i, j).Range.ParagraphFormat.Alignment='wdAlignParagraphLeft';
        DTI.Cell(i, j).VerticalAlignment='wdCellAlignVerticalCenter';
    end
end

  DTI.Cell(2, 1).Range.Text = '当事人或代理人签名:';
  DTI.Cell(3, 1).Range.Text = '勘验(检查)人签名:';

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  text_temp='备注： ';
  set(selection, 'Text',text_temp);
  selection.Font.Size=15;%设置文本字体大小
  selection.Font.Bold=0;%设置文本粗体
  
    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  
  %% 第八页开始
  
   end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  
  text_temp='粘   贴   专   页';
  set(selection, 'Text',text_temp);


  selection.Font.Size=16;%设置文本字体大小
  selection.Font.Bold=4;%设置文本粗体
  set(paragraphformat, 'LineSpacing',23);
  set(paragraphformat, 'Alignment','wdAlignParagraphCenter');

  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  
  %开始画第八页的图
  selection.Font.Bold=0;%设置文本粗体
  selection.Font.underline = 0;
  selection.Font.Size=15;%设置文本字体大小
  Tables=document.Tables.Add(selection.Range,1,1);
 
%设置边框
DTI=document.Tables.Item(7);
DTI.Borders.OutsideLineStyle='wdLineStyleSingle';
% DTI.Borders.OutsideLineWidth='wdLineWidth150pt';
DTI.Borders.InsideLineStyle='wdLineStyleSingle';
% DTI.Borders.InsideLineWidth='wdLineWidth150pt';
DTI.Rows.Alignment='wdAlignRowCenter';

column_width=510;
row_height=650 ;

  DTI.Columns.Item(1).Width =column_width;
  DTI.Rows.Item(1).Height =row_height;
  DTI.Cell(1, 1).Range.ParagraphFormat.Alignment='wdAlignParagraphCenter';
  DTI.Cell(1, 1).VerticalAlignment='wdCellAlignVerticalTop';
  DTI.Cell(1, 1).Range.Text = '检  测  磅  单';

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  
 %% 第九页开始
  
   end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  
  text_temp='粘   贴   专   页';
  set(selection, 'Text',text_temp);


  selection.Font.Size=16;%设置文本字体大小
  selection.Font.Bold=4;%设置文本粗体
  set(paragraphformat, 'LineSpacing',23);
  set(paragraphformat, 'Alignment','wdAlignParagraphCenter');

  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  
  %开始画第八页的图
  selection.Font.Bold=0;%设置文本粗体
  selection.Font.underline = 0;
  selection.Font.Size=15;%设置文本字体大小
  Tables=document.Tables.Add(selection.Range,2,1);
 
%设置边框
DTI=document.Tables.Item(8);
DTI.Borders.OutsideLineStyle='wdLineStyleSingle';
% DTI.Borders.OutsideLineWidth='wdLineWidth150pt';
DTI.Borders.InsideLineStyle='wdLineStyleSingle';
% DTI.Borders.InsideLineWidth='wdLineWidth150pt';
DTI.Rows.Alignment='wdAlignRowCenter';

column_width=510;
row_height=[325,325 ];

    DTI.Columns.Item(1).Width =column_width;

for i=1:2
    DTI.Rows.Item(i).Height =row_height(i);
end
for i = 1 : 2
    for j = 1 : 1
        DTI.Cell(i, j).Range.ParagraphFormat.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i, j).VerticalAlignment='wdCellAlignVerticalTop';
    end
end

  DTI.Cell(1, 1).Range.Text = '驾驶员身份证明复印件';
  DTI.Cell(2, 1).Range.Text = '车辆身份证明复印件';

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  
  
function edit1_Callback(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit1 as text
%        str2double(get(hObject,'String')) returns contents of edit1 as a double


% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit2_Callback(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit2 as text
%        str2double(get(hObject,'String')) returns contents of edit2 as a double


% --- Executes during object creation, after setting all properties.
function edit2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit3_Callback(hObject, eventdata, handles)
% hObject    handle to edit3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit3 as text
%        str2double(get(hObject,'String')) returns contents of edit3 as a double


% --- Executes during object creation, after setting all properties.
function edit3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit4_Callback(hObject, eventdata, handles)
% hObject    handle to edit4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit4 as text
%        str2double(get(hObject,'String')) returns contents of edit4 as a double


% --- Executes during object creation, after setting all properties.
function edit4_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit5_Callback(hObject, eventdata, handles)
% hObject    handle to edit5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit5 as text
%        str2double(get(hObject,'String')) returns contents of edit5 as a double


% --- Executes during object creation, after setting all properties.
function edit5_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit6_Callback(hObject, eventdata, handles)
% hObject    handle to edit6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit6 as text
%        str2double(get(hObject,'String')) returns contents of edit6 as a double


% --- Executes during object creation, after setting all properties.
function edit6_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit7_Callback(hObject, eventdata, handles)
% hObject    handle to edit7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit7 as text
%        str2double(get(hObject,'String')) returns contents of edit7 as a double


% --- Executes during object creation, after setting all properties.
function edit7_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit8_Callback(hObject, eventdata, handles)
% hObject    handle to edit8 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit8 as text
%        str2double(get(hObject,'String')) returns contents of edit8 as a double


% --- Executes during object creation, after setting all properties.
function edit8_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit8 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit9_Callback(hObject, eventdata, handles)
% hObject    handle to edit9 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit9 as text
%        str2double(get(hObject,'String')) returns contents of edit9 as a double


% --- Executes during object creation, after setting all properties.
function edit9_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit9 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit10_Callback(hObject, eventdata, handles)
% hObject    handle to edit10 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit10 as text
%        str2double(get(hObject,'String')) returns contents of edit10 as a double


% --- Executes during object creation, after setting all properties.
function edit10_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit10 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit11_Callback(hObject, eventdata, handles)
% hObject    handle to edit11 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit11 as text
%        str2double(get(hObject,'String')) returns contents of edit11 as a double


% --- Executes during object creation, after setting all properties.
function edit11_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit11 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit12_Callback(hObject, eventdata, handles)
% hObject    handle to edit12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit12 as text
%        str2double(get(hObject,'String')) returns contents of edit12 as a double


% --- Executes during object creation, after setting all properties.
function edit12_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit13_Callback(hObject, eventdata, handles)
% hObject    handle to edit13 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit13 as text
%        str2double(get(hObject,'String')) returns contents of edit13 as a double


% --- Executes during object creation, after setting all properties.
function edit13_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit13 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit14_Callback(hObject, eventdata, handles)
% hObject    handle to edit14 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit14 as text
%        str2double(get(hObject,'String')) returns contents of edit14 as a double


% --- Executes during object creation, after setting all properties.
function edit14_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit14 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit15_Callback(hObject, eventdata, handles)
% hObject    handle to edit15 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit15 as text
%        str2double(get(hObject,'String')) returns contents of edit15 as a double


% --- Executes during object creation, after setting all properties.
function edit15_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit15 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit16_Callback(hObject, eventdata, handles)
% hObject    handle to edit16 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit16 as text
%        str2double(get(hObject,'String')) returns contents of edit16 as a double


% --- Executes during object creation, after setting all properties.
function edit16_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit16 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit17_Callback(hObject, eventdata, handles)
% hObject    handle to edit17 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit17 as text
%        str2double(get(hObject,'String')) returns contents of edit17 as a double


% --- Executes during object creation, after setting all properties.
function edit17_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit17 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit18_Callback(hObject, eventdata, handles)
% hObject    handle to edit18 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit18 as text
%        str2double(get(hObject,'String')) returns contents of edit18 as a double


% --- Executes during object creation, after setting all properties.
function edit18_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit18 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit19_Callback(hObject, eventdata, handles)
% hObject    handle to edit19 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit19 as text
%        str2double(get(hObject,'String')) returns contents of edit19 as a double


% --- Executes during object creation, after setting all properties.
function edit19_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit19 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit20_Callback(hObject, eventdata, handles)
% hObject    handle to edit20 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit20 as text
%        str2double(get(hObject,'String')) returns contents of edit20 as a double


% --- Executes during object creation, after setting all properties.
function edit20_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit20 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit21_Callback(hObject, eventdata, handles)
% hObject    handle to edit21 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit21 as text
%        str2double(get(hObject,'String')) returns contents of edit21 as a double


% --- Executes during object creation, after setting all properties.
function edit21_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit21 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit22_Callback(hObject, eventdata, handles)
% hObject    handle to edit22 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit22 as text
%        str2double(get(hObject,'String')) returns contents of edit22 as a double


% --- Executes during object creation, after setting all properties.
function edit22_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit22 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit23_Callback(hObject, eventdata, handles)
% hObject    handle to edit23 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit23 as text
%        str2double(get(hObject,'String')) returns contents of edit23 as a double


% --- Executes during object creation, after setting all properties.
function edit23_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit23 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit24_Callback(hObject, eventdata, handles)
% hObject    handle to edit24 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit24 as text
%        str2double(get(hObject,'String')) returns contents of edit24 as a double


% --- Executes during object creation, after setting all properties.
function edit24_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit24 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit25_Callback(hObject, eventdata, handles)
% hObject    handle to edit25 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit25 as text
%        str2double(get(hObject,'String')) returns contents of edit25 as a double


% --- Executes during object creation, after setting all properties.
function edit25_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit25 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit26_Callback(hObject, eventdata, handles)
% hObject    handle to edit26 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit26 as text
%        str2double(get(hObject,'String')) returns contents of edit26 as a double


% --- Executes during object creation, after setting all properties.
function edit26_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit26 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit27_Callback(hObject, eventdata, handles)
% hObject    handle to edit27 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit27 as text
%        str2double(get(hObject,'String')) returns contents of edit27 as a double


% --- Executes during object creation, after setting all properties.
function edit27_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit27 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit28_Callback(hObject, eventdata, handles)
% hObject    handle to edit28 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit28 as text
%        str2double(get(hObject,'String')) returns contents of edit28 as a double


% --- Executes during object creation, after setting all properties.
function edit28_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit28 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit30_Callback(hObject, eventdata, handles)
% hObject    handle to edit30 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit30 as text
%        str2double(get(hObject,'String')) returns contents of edit30 as a double


% --- Executes during object creation, after setting all properties.
function edit30_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit30 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit31_Callback(hObject, eventdata, handles)
% hObject    handle to edit31 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit31 as text
%        str2double(get(hObject,'String')) returns contents of edit31 as a double


% --- Executes during object creation, after setting all properties.
function edit31_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit31 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit32_Callback(hObject, eventdata, handles)
% hObject    handle to edit32 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit32 as text
%        str2double(get(hObject,'String')) returns contents of edit32 as a double


% --- Executes during object creation, after setting all properties.
function edit32_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit32 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit33_Callback(hObject, eventdata, handles)
% hObject    handle to edit33 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit33 as text
%        str2double(get(hObject,'String')) returns contents of edit33 as a double


% --- Executes during object creation, after setting all properties.
function edit33_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit33 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit34_Callback(hObject, eventdata, handles)
% hObject    handle to edit34 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit34 as text
%        str2double(get(hObject,'String')) returns contents of edit34 as a double


% --- Executes during object creation, after setting all properties.
function edit34_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit34 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit35_Callback(hObject, eventdata, handles)
% hObject    handle to edit35 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit35 as text
%        str2double(get(hObject,'String')) returns contents of edit35 as a double


% --- Executes during object creation, after setting all properties.
function edit35_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit35 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit36_Callback(hObject, eventdata, handles)
% hObject    handle to edit36 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit36 as text
%        str2double(get(hObject,'String')) returns contents of edit36 as a double


% --- Executes during object creation, after setting all properties.
function edit36_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit36 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit37_Callback(hObject, eventdata, handles)
% hObject    handle to edit37 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit37 as text
%        str2double(get(hObject,'String')) returns contents of edit37 as a double


% --- Executes during object creation, after setting all properties.
function edit37_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit37 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit38_Callback(hObject, eventdata, handles)
% hObject    handle to edit38 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit38 as text
%        str2double(get(hObject,'String')) returns contents of edit38 as a double


% --- Executes during object creation, after setting all properties.
function edit38_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit38 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit39_Callback(hObject, eventdata, handles)
% hObject    handle to edit39 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit39 as text
%        str2double(get(hObject,'String')) returns contents of edit39 as a double


% --- Executes during object creation, after setting all properties.
function edit39_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit39 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit40_Callback(hObject, eventdata, handles)
% hObject    handle to edit40 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit40 as text
%        str2double(get(hObject,'String')) returns contents of edit40 as a double


% --- Executes during object creation, after setting all properties.
function edit40_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit40 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit42_Callback(hObject, eventdata, handles)
% hObject    handle to edit42 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit42 as text
%        str2double(get(hObject,'String')) returns contents of edit42 as a double


% --- Executes during object creation, after setting all properties.
function edit42_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit42 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit43_Callback(hObject, eventdata, handles)
% hObject    handle to edit43 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit43 as text
%        str2double(get(hObject,'String')) returns contents of edit43 as a double


% --- Executes during object creation, after setting all properties.
function edit43_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit43 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit44_Callback(hObject, eventdata, handles)
% hObject    handle to edit44 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit44 as text
%        str2double(get(hObject,'String')) returns contents of edit44 as a double


% --- Executes during object creation, after setting all properties.
function edit44_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit44 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit45_Callback(hObject, eventdata, handles)
% hObject    handle to edit45 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit45 as text
%        str2double(get(hObject,'String')) returns contents of edit45 as a double


% --- Executes during object creation, after setting all properties.
function edit45_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit45 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit46_Callback(hObject, eventdata, handles)
% hObject    handle to edit46 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit46 as text
%        str2double(get(hObject,'String')) returns contents of edit46 as a double


% --- Executes during object creation, after setting all properties.
function edit46_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit46 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit47_Callback(hObject, eventdata, handles)
% hObject    handle to edit47 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit47 as text
%        str2double(get(hObject,'String')) returns contents of edit47 as a double


% --- Executes during object creation, after setting all properties.
function edit47_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit47 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit48_Callback(hObject, eventdata, handles)
% hObject    handle to edit48 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit48 as text
%        str2double(get(hObject,'String')) returns contents of edit48 as a double


% --- Executes during object creation, after setting all properties.
function edit48_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit48 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit49_Callback(hObject, eventdata, handles)
% hObject    handle to edit49 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit49 as text
%        str2double(get(hObject,'String')) returns contents of edit49 as a double


% --- Executes during object creation, after setting all properties.
function edit49_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit49 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit50_Callback(hObject, eventdata, handles)
% hObject    handle to edit50 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit50 as text
%        str2double(get(hObject,'String')) returns contents of edit50 as a double


% --- Executes during object creation, after setting all properties.
function edit50_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit50 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit51_Callback(hObject, eventdata, handles)
% hObject    handle to edit51 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit51 as text
%        str2double(get(hObject,'String')) returns contents of edit51 as a double


% --- Executes during object creation, after setting all properties.
function edit51_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit51 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit52_Callback(hObject, eventdata, handles)
% hObject    handle to edit52 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit52 as text
%        str2double(get(hObject,'String')) returns contents of edit52 as a double


% --- Executes during object creation, after setting all properties.
function edit52_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit52 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

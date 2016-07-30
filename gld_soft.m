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
%% �Ȼ�ȡ��������
%��������Ϣ��ȡ�����񲿷�
xingmin = (get(handles.edit1,'string'));%����������
xingbie = (get(handles.edit2,'string'));%�Ա�
shenfenzhenghao = (get(handles.edit3,'string'));%���֤��
zhuzhi = (get(handles.edit4,'string'));%סַ
zhiye = (get(handles.edit5,'string'));%ְҵ
nianling = (get(handles.edit30,'string'));%����
gongmindianhua = (get(handles.edit31,'string'));%�����˵绰
anjianguanxi = (get(handles.edit39,'string'));%������ϵ
danweizhiwu = (get(handles.edit40,'string'));%��λְ��

%��������Ϣ��ȡ�����˲���
mingcheng = (get(handles.edit6,'string'));%��������
fadingdaibiaoren = (get(handles.edit7,'string'));%���˷���������
dizhi = (get(handles.edit8,'string'));%��ַ
farendianhua = (get(handles.edit32,'string'));%���˵绰


% Υ��ʱ����ȡ
weifanian = (get(handles.edit9,'string'));
weifayue = (get(handles.edit10,'string'));
weifari = (get(handles.edit11,'string'));
weifashi = (get(handles.edit12,'string'));
weifafen = (get(handles.edit13,'string'));

%Υ��·��
shendao = (get(handles.edit49,'string'));%ʡ��
chuzi = (get(handles.edit50,'string'));%����
fangxiangxiang = (get(handles.edit51,'string'));%������
fangxiangxingshi = (get(handles.edit52,'string'));%������ʻ

%������Ϣ��ȡ
chepaihao = (get(handles.edit14,'string'));%���ƺ�
zhoushuliang = (get(handles.edit15,'string'));%������
huocheleixing = (get(handles.edit16,'string'));%��������
huowuleixing = (get(handles.edit17,'string'));%��������
chehuozongliang = (get(handles.edit18,'string'));%��������
chaoxiandunshu = (get(handles.edit19,'string'));%���޶���
jiashiyuanzhengjianleixing  = (get(handles.edit20,'string'));%��ʻԱ֤������
cheliangzhengjianleixing  = (get(handles.edit21,'string'));%����֤������
xingzhengchufaleixing  = (get(handles.edit22,'string'));%������������
yinhangzhanghao  = (get(handles.edit23,'string'));%�����˺�

% ����ʱ����ȡ
chufanian = (get(handles.edit24,'string'));
chufayue = (get(handles.edit25,'string'));
chufari = (get(handles.edit26,'string'));
chufashi = (get(handles.edit27,'string'));
chufafen = (get(handles.edit28,'string'));

% ִ����Ա��Ϣ��ȡ
zhifadidian = (get(handles.edit33,'string'));%ִ���ص�
zhifarenyuan1 = (get(handles.edit34,'string'));%ִ����Ա1
zhifarenyuan2 = (get(handles.edit35,'string'));%ִ����Ա2
jiluren = (get(handles.edit36,'string'));%ִ����Ա2
zhifazhenghao1 = (get(handles.edit37,'string'));%ִ��֤��1
zhifazhenghao2 = (get(handles.edit38,'string'));%ִ��֤��2
zhifanian = (get(handles.edit42,'string'));%ִ��ʱ��
zhifayue = (get(handles.edit43,'string'));
zhifari = (get(handles.edit44,'string'));
zhifashi1 = (get(handles.edit45,'string'));
zhifafen1 = (get(handles.edit46,'string'));
zhifashi2 = (get(handles.edit47,'string'));
zhifafen2 = (get(handles.edit48,'string'));




%��Matlab����Word�ĵ�
%��Matlab����һ�γ��򣬿�������Word�ĵ����ĵ��к��б�񣬴������£�
filespec = ['�Զ����Ա���' datestr(now,30) '.doc'];
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
%ҳ������
document.PageSetup.TopMargin = 60;
document.PageSetup.BottomMargin = 45;
document.PageSetup.LeftMargin = 40;
document.PageSetup.RightMargin = 40;



% ��һҳ
set(content, 'Start',0);
title='��������������';
set(content, 'Text',title);
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
rr=document.Range(0,7);%ѡ���ı�
rr.Font.Size=16;%�����ı������С
rr.Font.Bold=4;%�����ı�����
rr.Font.Name = '����';%ѡ������
% rr.Font.Bold = 'wdToggle';
end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
selection.Font.Size=10;
selection.MoveDown;
selection.TypeParagraph;
 set(paragraphformat, 'Alignment','wdAlignParagraphRight');
 
 text1='����·��  ����    ��     ��';
set(selection, 'Text',text1);
dd=document.Range(8,14);%ѡ���ı�
 dd.Font.underline = 2;
 dd.Font.Size=12;%�����ı������С
dd.Font.Bold=4;%�����ı�����

dd=document.Range(15,28);%ѡ���ı�
%  dd.Font.underline = 2;
 dd.Font.Size=15;%�����ı������С
% dd.Font.Bold=4;%�����ı�����

% ��ο�����ʱ�س�����˼
end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 selection.TypeParagraph;
 
selection.Font.Size=15;
Tables=document.Tables.Add(selection.Range,4,8);
 
%���ñ߿�
DTI=document.Tables.Item(1);
DTI.Borders.OutsideLineStyle='wdLineStyleSingle';
% DTI.Borders.OutsideLineWidth='wdLineWidth150pt';
DTI.Borders.InsideLineStyle='wdLineStyleSingle';
% DTI.Borders.InsideLineWidth='wdLineWidth150pt';
DTI.Rows.Alignment='wdAlignRowCenter';
% end_of_doc = get(content,'end');
% set(selection,'Start',end_of_doc);
% selection.TypeParagraph;
% set(selection, 'Text','����ǩ�֣�            ��    ��    ��');
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
  
  DTI.Cell(1, 1).Range.Text = '������';
  DTI.Cell(1, 2).Range.Text = '����';
  DTI.Cell(1, 3).Range.Text = '����';
  DTI.Cell(1, 4).Range.Text = xingmin;
    DTI.Cell(1, 5).Range.Text = '�Ա�';
   DTI.Cell(1, 6).Range.Text = xingbie;
  DTI.Cell(1, 7).Range.Text = '���֤��';
  DTI.Cell(1, 8).Range.Text = shenfenzhenghao;
   DTI.Cell(1, 8).Range.Font.Size = 12 ;%���õ�Ԫ��������С
  DTI.Cell(2, 3).Range.Text = 'סַ';
  DTI.Cell(2, 4).Range.Text = zhuzhi;
  DTI.Cell(2, 4).Range.Font.Size = 12 ;
  DTI.Cell(2, 5).Range.Text = 'ְҵ';
  DTI.Cell(2, 6).Range.Text = zhiye;
  DTI.Cell(2, 6).Range.Font.Size = 12 ;
  DTI.Cell(3, 2).Range.Text = '���˻�������֯';
  DTI.Cell(3, 3).Range.Text = '����';
  DTI.Cell(3, 4).Range.Text = mingcheng;
  DTI.Cell(3, 4).Range.Font.Size = 12 ;
  DTI.Cell(4, 3).Range.Text = '��ַ';
  DTI.Cell(4, 4).Range.Text = dizhi;
  DTI.Cell(4, 4).Range.Font.Size = 12 ;
  DTI.Cell(3, 5).Range.Text = '����������';
  DTI.Cell(3, 6).Range.Text = fadingdaibiaoren;
  
  % ��ο�����ʱ�س�����˼
end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
 selection.TypeParagraph;
  set(paragraphformat, 'LineSpacingRule','wdLineSpaceExactly');
 set(paragraphformat, 'LineSpacing',23);


text2='    Υ����ʵ��֤�ݣ�';%
set(selection, 'Text',text2);
selection.Font.Size=15;
end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 
text3=['������'    xingmin      '��' weifanian '��' weifayue '��' weifari  '��'  weifashi  'ʱ' weifafen   ...
      '�֣���ʻ�ƺ�Ϊ' chepaihao '��' zhoushuliang  ' �� ' huocheleixing  '����'  huowuleixing ...
      '�� ' shendao  'ʡ������ʻ���������й�·�����·����Ա��⣬��������' chehuozongliang '�֣����� ' chaoxiandunshu ...
      '�֣������Գ���������ʻ��·��ǰ����ʵ�ɡ��ֳ���¼������ѯ�ʱ�¼������������������ ��ʻԱ' jiashiyuanzhengjianleixing ...
      '  ��ӡ������������  '    cheliangzhengjianleixing   ' ��ӡ����������֤'  xingmin  ' ��Υ����ʵ���Գ��ϡ�'];%
set(selection, 'Text',text3);
selection.Font.Size=13;
selection.Font.Bold=4;%�����ı�����
set(paragraphformat, 'Alignment','wdAlignParagraphJustify');
selection.Font.underline = 2;
  set(paragraphformat, 'LineSpacingRule','wdLineSpaceExactly');
 set(paragraphformat, 'LineSpacing',23);
 
end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 selection.TypeParagraph;
 text4=['    ������ʵΥ���ˡ��л����񹲺͹���·��������ʮ���� ������ʮ��������·��ȫ��������������ʮ������һ�� �Ĺ涨������ ���л����񹲺͹���·��������ʮ�������������·��ȫ��������������ʮ�����涨���������� ' xingzhengchufaleixing ' ������������'];%
set(selection, 'Text',text4);
selection.Font.Size=13;
selection.Font.Bold=4;%�����ı�����
  set(paragraphformat, 'LineSpacingRule','wdLineSpaceExactly');
 set(paragraphformat, 'LineSpacing',23);


end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 selection.TypeParagraph;
 text5=['    ���Է���ģ��������յ���������֮����15���ڽ��� �����н�������Ӫҵ��   ���˺�Ϊ��' yinhangzhanghao '���ڲ��ɵ�����ÿ�հ����������3%�Ӵ����'];%
set(selection, 'Text',text5);
selection.Font.Size=13;
selection.Font.Bold=4;%�����ı�����

end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 selection.TypeParagraph;
 text6=['    �����������������������������60������   �����н�ͨ�����  �����������飬������������������������Ժ�����������ϣ�����������ִֹͣ�У��������й涨�ĳ��⡣���ڲ������������顢���������������ֲ����еģ������ؽ�������������Ժǿ��ִ�л��������йع涨ǿ��ִ�С�'];%
set(selection, 'Text',text6);
selection.Font.Size=13;
selection.Font.Bold=4;%�����ı�����

end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 selection.TypeParagraph;
 selection.TypeParagraph;
text7=[ chufanian '   ��' chufayue '��'  chufari  '     ��' ];%
set(selection, 'Text',text7);
selection.Font.Size=14;
selection.Font.Name = '����';%ѡ������
 selection.Font.Bold=0;%�����ı�����
set(paragraphformat, 'Alignment','wdAlignParagraphRight');

end_of_doc = get(content,'end');
 set(selection,'Start',end_of_doc);
 selection.MoveDown;
 selection.TypeParagraph;
 selection.TypeParagraph;
text8=[ '��������һʽ���ݣ�һ�ݴ����һ�ݽ������˻�������ˡ���' ];%
set(selection, 'Text',text8);
selection.Font.Size=14;
selection.Font.Name = '����';%ѡ������
 selection.Font.Bold=0;%�����ı�����
set(paragraphformat, 'Alignment','wdAlignParagraphLeft');

%�ڶ�ҳ
selection.MoveDown;
selection.TypeParagraph;
selection.TypeParagraph;
selection.TypeParagraph;
 text21='��  ��  ��  ��  ��';
set(selection, 'Text',text21);


 selection.Font.Size=16;%�����ı������С
selection.Font.Bold=4;%�����ı�����
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');

selection.MoveDown;
selection.TypeParagraph;
 set(paragraphformat, 'Alignment','wdAlignParagraphRight');
 
 text_temp='����·��';%  ����    ��     ��';
set(selection, 'Text',text_temp);

 selection.Font.underline = 2;
 selection.Font.Size=12;%�����ı������С
 selection.Font.Bold=4;%�����ı�����

end_of_doc = get(content,'end');%���λ���ƶ������
 set(selection,'Start',end_of_doc);

 text_temp='  ����    ��     ��';
 set(selection, 'Text',text_temp);
%�µ����ֵĸ�ʽ��Ҫ�������ã������������֮ǰ�ĸ�ʽ
 selection.Font.Bold=0;%�����ı�����
 selection.Font.underline = 0;
 selection.Font.Size=15;%�����ı������С
 end_of_doc = get(content,'end');%���λ���ƶ������,ÿ������������Ϻ�Ӧ�����������
 set(selection,'Start',end_of_doc);



%��ʼ���ڶ�ҳ��ͼ
 selection.Font.Bold=0;%�����ı�����
 selection.Font.underline = 0;
 selection.Font.Size=15;%�����ı������С
 Tables=document.Tables.Add(selection.Range,10,8);
 
%���ñ߿�
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
 
  DTI.Cell(1, 1).Range.Text = '������Դ';
  DTI.Cell(1, 2).Range.Text = 'ִ����鷢��';
  DTI.Cell(1, 2).Range.Font.Size = 13 ;%���õ�Ԫ��������С
  DTI.Cell(1, 2).Range.Font.Bold = 2 ;%���õ�Ԫ��������С
  DTI.Cell(1, 3).Range.Text = '�ܰ�ʱ��';
  DTI.Cell(1, 4).Range.Text = [weifanian '��' weifayue '��' weifari '��' weifashi 'ʱ' weifafen '��' ];
  DTI.Cell(2, 1).Range.Text = '����';
  DTI.Cell(2, 2).Range.Text = '�������Գ���������ʻ��·��';
  DTI.Cell(2, 2).Range.Font.Size = 13 ;%���õ�Ԫ��������С
  DTI.Cell(2, 2).Range.Font.Bold = 2 ;%���õ�Ԫ��������С
  DTI.Cell(3, 1).Range.Text = '�����˻������';
  DTI.Cell(3, 2).Range.Text = '����';
  DTI.Cell(3, 3).Range.Text = '����';
  DTI.Cell(3, 4).Range.Text = xingmin;
  DTI.Cell(3, 5).Range.Text = '�Ա�';
  DTI.Cell(3, 6).Range.Text = xingbie;
  DTI.Cell(3, 7).Range.Text = '����';
  DTI.Cell(3, 8).Range.Text = nianling;
  DTI.Cell(4, 3).Range.Text = 'סַ';
  DTI.Cell(4, 4).Range.Text = zhuzhi;
  DTI.Cell(4, 4).Range.Font.Size = 12 ;%���õ�Ԫ��������С
  DTI.Cell(4, 5).Range.Text = '���֤��';
  DTI.Cell(4, 6).Range.Text = shenfenzhenghao;
  DTI.Cell(4, 6).Range.Font.Size = 12 ;%���õ�Ԫ��������С
  DTI.Cell(4, 7).Range.Text = '��ϵ�绰';
  DTI.Cell(4, 8).Range.Text = gongmindianhua;
  DTI.Cell(5, 2).Range.Text = '���˻�������֯';
  DTI.Cell(5, 3).Range.Text = '����';
  DTI.Cell(5, 4).Range.Text = mingcheng;
  DTI.Cell(5, 5).Range.Text = '����������';
  DTI.Cell(5, 6).Range.Text = fadingdaibiaoren;
  DTI.Cell(6, 3).Range.Text = '��ַ';
  DTI.Cell(6, 4).Range.Text = dizhi;
  DTI.Cell(6, 4).Range.Font.Size = 12 ;%���õ�Ԫ��������С
  DTI.Cell(6, 5).Range.Text = '��ϵ�绰';
  DTI.Cell(6, 6).Range.Text = farendianhua;
  DTI.Cell(7, 1).Range.Text = '�����������';
  DTI.Cell(7, 2).Range.Text = [weifanian '��' weifayue '��' weifari '��' weifashi 'ʱ' weifafen '��' ...
                                '�������й�·�����·��ִ����Ա' zhifarenyuan1  '�� ' zhifarenyuan2 ...
                                '  �����м���з���, ��ʻԱ' xingmin '��ʻ���ƺ�Ϊ ' chepaihao      ...
                                ' �� ' zhoushuliang   ' ��' huocheleixing    '�������� ' huowuleixing    ...
                                ' ,��' shendao  'ʡ��' chuzi  '����' fangxiangxiang  '������ ' fangxiangxingshi   ...
                                '   ������ʻ����������⣬�ó��������أ��������ߣ� '  chehuozongliang   ...
                                '  �֣��ף����������Գ���������ʻ��·��'];
  DTI.Cell(7, 2).Range.Font.Size = 12 ;%���õ�Ԫ��������С
  DTI.Cell(8, 1).Range.Text = '��������';
  DTI.Cell(8, 2).Range.Text = '���л����񹲺͹���·��������ʮ���� ������ʮ��������·��ȫ��������������ʮ������һ�';
  DTI.Cell(8, 2).Range.Font.Size = 13 ;%���õ�Ԫ��������С
  DTI.Cell(8, 2).Range.Font.Bold = 2 ;%���õ�Ԫ��������С
  DTI.Cell(8, 3).Range.Text = '�ܰ��������';
  DTI.Cell(8, 4).Range.Text = '                                    ǩ����               ʱ�䣺    �� �� ��';
  DTI.Cell(8, 4).Range.ParagraphFormat.Alignment='wdAlignParagraphLeft';
  DTI.Cell(8, 4).Range.Font.Size = 13 ;%���õ�Ԫ��������С
  DTI.Cell(8, 4).VerticalAlignment='wdCellAlignVerticalBottom';
  DTI.Cell(9, 1).Range.Text = '�������������';
  DTI.Cell(9, 2).Range.Text = '                                      ǩ����                  .        ʱ�䣺      ��   ��    ��';
  DTI.Cell(9, 2).Range.ParagraphFormat.Alignment='wdAlignParagraphRight';
  DTI.Cell(9, 2).VerticalAlignment='wdCellAlignVerticalBottom';
  DTI.Cell(9, 2).Range.Font.Size = 13 ;%���õ�Ԫ��������С
  DTI.Cell(10, 1).Range.Text = '��ע';
  DTI.Cell(10, 2).Range.Text = '����������';
  DTI.Cell(10, 2).Range.Font.Size = 13 ;%���õ�Ԫ��������С
  DTI.Cell(10, 2).Range.Font.Bold = 2 ;%���õ�Ԫ��������С
  % ������Ҫ�޸ĸ�ʽ����յ�����
  
  
  %����ҳ
   end_of_doc = get(content,'end');%���λ���ƶ������,ÿ������������Ϻ�Ӧ�����������
 set(selection,'Start',end_of_doc);
  selection.MoveDown;
selection.TypeParagraph;
selection.TypeParagraph;
selection.TypeParagraph;
 text_temp='��  ��  ��  ¼';
set(selection, 'Text',text_temp);


 selection.Font.Size=16;%�����ı������С
selection.Font.Bold=4;%�����ı�����
set(paragraphformat, 'Alignment','wdAlignParagraphCenter');

selection.MoveDown;
selection.TypeParagraph;
 
 end_of_doc = get(content,'end');%���λ���ƶ������,ÿ������������Ϻ�Ӧ�����������
 set(selection,'Start',end_of_doc);
  
  %��ʼ������ҳ��ͼ
 selection.Font.Bold=0;%�����ı�����
 selection.Font.underline = 0;
 selection.Font.Size=15;%�����ı������С
 Tables=document.Tables.Add(selection.Range,11,7);
 
%���ñ߿�
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
  
  DTI.Cell(1, 1).Range.Text = 'ִ���ص�';
  DTI.Cell(1, 2).Range.Text = zhifadidian ;
  DTI.Cell(1, 3).Range.Text = 'ִ��ʱ��';
  DTI.Cell(1, 4).Range.Text = [zhifanian '��' zhifayue  '��' zhifari '��'...
                               zhifashi1 'ʱ'  zhifafen1 '����' zhifashi2 'ʱ' zhifafen2 '��' ];
  DTI.Cell(2, 1).Range.Text = 'ִ����Ա';
  DTI.Cell(2, 2).Range.Text = zhifarenyuan1 ;
  DTI.Cell(3, 2).Range.Text = zhifarenyuan2 ;
  DTI.Cell(2, 3).Range.Text = 'ִ��֤��';
  DTI.Cell(2, 4).Range.Text = zhifazhenghao1 ;
  DTI.Cell(3, 4).Range.Text = zhifazhenghao2 ;
  DTI.Cell(2, 5).Range.Text = '��¼��';
  DTI.Cell(2, 6).Range.Text = jiluren ;
  DTI.Cell(4, 1).Range.Text = '�ֳ���Ա�������';
  DTI.Cell(4, 2).Range.Text = '��   ��';
  DTI.Cell(4, 3).Range.Text =  xingmin ;
  DTI.Cell(4, 4).Range.Text = '��   ��';
  DTI.Cell(4, 5).Range.Text =  xingbie;
  DTI.Cell(5, 2).Range.Text = '���֤��';
  DTI.Cell(5, 3).Range.Text = shenfenzhenghao ;
  DTI.Cell(5, 4).Range.Text = '�밸����ϵ';
  DTI.Cell(5, 5).Range.Text = anjianguanxi ;
  DTI.Cell(6, 2).Range.Text = '��λ��ְ��';
  DTI.Cell(6, 3).Range.Text = danweizhiwu ;
  DTI.Cell(6, 4).Range.Text = '��ϵ�绰';
  DTI.Cell(6, 5).Range.Text = gongmindianhua ;
  DTI.Cell(7, 2).Range.Text = '��ϵ��ַ';
  DTI.Cell(7, 3).Range.Text = zhuzhi ;
  DTI.Cell(8, 2).Range.Text = '����������';
  DTI.Cell(8, 3).Range.Text = chepaihao ;
  DTI.Cell(8, 4).Range.Text = '����������';
  DTI.Cell(8, 5).Range.Text = huocheleixing ;
  DTI.Cell(9, 1).Range.Text = '��Ҫ����';
  DTI.Cell(9, 2).Range.Text = ['�ڼ���з��֣���ʻԱ'  xingmin  '��  '...
                                weifanian  ' �� '  weifayue  ' �� '  weifari ...
                                ' �� ' weifashi   ' ʱ ' weifafen  ...
                                ' �֣���ʻ '     ' �ƺ�Ϊ '  chepaihao ...
                                '�� ' zhoushuliang  ' �� ' huocheleixing '  �������� ' huowuleixing ...
                                ' �� ' shendao  ' ʡ�� ' chuzi  ' ���� ' fangxiangxiang  '������ ' fangxiangxingshi  ' ������ʻ�������й�·�����·����Ա������������˼�⣬����⣬�������� '...
                                chehuozongliang  '  �֡�                                        �����������ݣ�    '...
                                '                                 __________________________________________________________________________ ������¼���ѿ����������������������������ʵ���󡣡�'...
                                 '�ֳ���Աǩ����             ʱ�䣺   ��  ��  ��'];
  DTI.Cell(10, 2).Range.Text = '��ע';
  DTI.Cell(10, 2).Range.ParagraphFormat.Alignment='wdAlignParagraphLeft';
  DTI.Cell(11, 2).Range.Text = 'ִ����Աǩ����                         ʱ�䣺   ��  ��  ��';
  DTI.Cell(11, 2).Range.ParagraphFormat.Alignment='wdAlignParagraphLeft';
  end_of_doc = get(content,'end');%���λ���ƶ������,ÿ������������Ϻ�Ӧ�����������
  set(selection,'Start',end_of_doc);
%   set(selection, 'Text',DTI.Cell(11, 2).Range.Text);
  selection.MoveDown;
  selection.TypeParagraph;
  
  %����ҳ��ʼ
   text_temp='ѯ   ��   ��   ¼';
   set(selection, 'Text',text_temp);
   selection.Font.Size=16;%�����ı������С
   selection.Font.Bold=4;%�����ı�����
   set(paragraphformat, 'Alignment','wdAlignParagraphCenter');
   
    end_of_doc = get(content,'end');%���λ���ƶ������,ÿ������������Ϻ�Ӧ�����������
    set(selection,'Start',end_of_doc);
    selection.MoveDown;
    selection.TypeParagraph;
  
    %��ʼ������ҳ��ͼ
 selection.Font.Bold=0;%�����ı�����
 selection.Font.underline = 0;
 selection.Font.Size=15;%�����ı������С
 Tables=document.Tables.Add(selection.Range,8,4);
 
%���ñ߿�
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
  
DTI.Cell(1, 1).Range.Text = 'ʱ��';
DTI.Cell(2, 1).Range.Text = 'ѯ�ʵص�';
DTI.Cell(3, 1).Range.Text = 'ѯ����';
DTI.Cell(3, 3).Range.Text = '��¼��';
DTI.Cell(4, 1).Range.Text = '��ѯ����';
DTI.Cell(4, 2).Range.Text = xingmin ;
DTI.Cell(4, 3).Range.Text = '�밸����ϵ';
DTI.Cell(4, 4).Range.Text = anjianguanxi ;
DTI.Cell(5, 1).Range.Text = '�Ա�';
DTI.Cell(5, 2).Range.Text = xingbie ;
DTI.Cell(5, 3).Range.Text = '����';
DTI.Cell(5, 4).Range.Text = nianling ;
DTI.Cell(6, 1).Range.Text = '���֤��';
DTI.Cell(6, 2).Range.Text = shenfenzhenghao ;
DTI.Cell(6, 3).Range.Text = '�绰';
DTI.Cell(6, 4).Range.Text = gongmindianhua ;
DTI.Cell(7, 1).Range.Text = '������λ��ְ��';
DTI.Cell(7, 2).Range.Text = danweizhiwu ;
DTI.Cell(8, 1).Range.Text = '��ϵ��ַ';
DTI.Cell(8, 2).Range.Text = zhuzhi ;
  
  end_of_doc = get(content,'end');%���λ���ƶ������,ÿ������������Ϻ�Ӧ�����������
  set(selection,'Start',end_of_doc);
   set(paragraphformat, 'LineSpacing',20);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['    �����ǵ����й�·�����·��ִ����Ա' zhifarenyuan1  ' ��' zhifarenyuan2  ' ��'...
        '�������ǵ�֤����ִ��֤������ֱ��� ' zhifazhenghao1  '  �� ' zhifazhenghao2  '  ��'...
        '����ȷ�ϡ�����������ѯ�ʣ�����ʵ�ش��������⣬ִ����Ա������ֱ��������ϵ�ģ����������رܡ�'];%
  set(selection, 'Text',text2);
  selection.Font.Size=15;
  
  %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ����Ƿ�����ִ����Ա�رܣ�                                         ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�𣺷�                                        '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı�����
  
    %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ����ʲô����?                                          ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['��' xingmin ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı�����
  
      %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ����������?                                            ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�� '  nianling  ' �� '   ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı�����
  
        %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ��������ס��ʲô�ط���                                            ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['��' zhuzhi  ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı�����
  
 %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ����ܳ�ʾһ�����֤����                                            ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�����֤����' shenfenzhenghao   ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı����� 
  
   %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ�����ʲô��λ������                                             ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�� ' danweizhiwu  ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı����� 
  
     %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�ʣ�' weifanian  ' ��  ' weifayue  ' �� ' weifari   ' �ճ��ƺ�Ϊ ' chepaihao  ' ������' shendao  'ʡ���Ͻ��л�������ʱ�����ʻ����  '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='���ǵ�';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı����� 
  
         %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ�������˭��  ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
    set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�� ' ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı����� 
  
      end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
   set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='��ѯ����ǩ����ʱ�䣺                         ѯ����ǩ����ʱ��';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%�����ı�����
  
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='____________________                       ______________________';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%�����ı�����
  
  

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='��ע������������';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%�����ı�����

  
    %����ҳ��ʼ
  
    
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='��ҳ';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%�����ı�����
  

  
  %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ������Ƕ��٣� ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�� ' chepaihao  ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı����� 
  
  %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ������������ʲô��� ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�� ' huowuleixing  ];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı����� 
  
  %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ�װ�˶��ٶ֣� ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�𣺷�   '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı�����
  
  %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ���ʲô�ط������ģ� ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�𣺷�                                        '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı�����
  
  %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ�Ŀ�ĵ���ʲô�ط��� ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�𣺷�                                        '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı�����
  
  %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ������˷��Ƕ��٣� ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�𣺷�                                        '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı�����
  
  %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ����Ǵ�����·��ʻ������ģ�  ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�𣺷�                                        '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı�����
  
  %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�ʣ������ǳ����Ǽ�⣬�������� ' chehuozongliang  '�֣��Դ˼�����������飿  '];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['�� ''û������'];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı�����
  
  %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ��������в������  ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['��''û��'];%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=0;%�����ı�����
  
  %�ʴ�����ģ��
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='�ʣ����ϼ�¼�Ƿ���ʵ�������ٿ�һ�£�����ʵ����ǩ�֡� ';%
  set(selection, 'Text',text2);
  selection.Font.Size=13; 
  selection.Font.Bold=4;%�����ı�����

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2=['��_____________________________________________________'...
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
  selection.Font.Bold=0;%�����ı�����
  
        end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
    selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
   set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='��ѯ����ǩ����ʱ�䣺                         ѯ����ǩ����ʱ��';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%�����ı�����

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='____________________                       ______________________';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%�����ı�����
  
  

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  text2='��ע������������';%
  set(selection, 'Text',text2);
  selection.Font.Size=15; 
  selection.Font.Bold=0;%�����ı�����
  
  %����ҳ��ʼ
  
    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  
  text_temp='���飨��飩��¼';
  set(selection, 'Text',text_temp);


  selection.Font.Size=16;%�����ı������С
  selection.Font.Bold=4;%�����ı�����
  set(paragraphformat, 'Alignment','wdAlignParagraphCenter');

  selection.MoveDown;
  selection.TypeParagraph;
  
  text_temp='���ɣ� ';
  set(selection, 'Text',text_temp);
  selection.Font.Size=15;%�����ı������С
  selection.Font.Bold=0;%�����ı�����
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  
  text_temp='_______________________�������Գ���������ʻ��·��__________________ ';
  set(selection, 'Text',text_temp);
  selection.Font.Size=13;%�����ı������С
  selection.Font.Bold=4;%�����ı�����
  selection.Font.underline = 2;
  
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  
    %��ʼ������ҳ��ͼ
  selection.Font.Bold=0;%�����ı�����
  selection.Font.underline = 0;
  selection.Font.Size=15;%�����ı������С
  Tables=document.Tables.Add(selection.Range,9,6);
 
%���ñ߿�
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

DTI.Cell(1, 1).Range.Text = '���飨��飩ʱ��';
DTI.Cell(2, 1).Range.Text = '���飨��飩����';
DTI.Cell(2, 3).Range.Text = '�������';
DTI.Cell(3, 1).Range.Text = '���飨��飩��';
DTI.Cell(3, 3).Range.Text = '��λ��ְ��';
DTI.Cell(3, 5).Range.Text = 'ִ��֤��';
DTI.Cell(4, 1).Range.Text = '���飨��飩��';
DTI.Cell(4, 3).Range.Text = '��λ��ְ��';
DTI.Cell(4, 5).Range.Text = 'ִ��֤��';
DTI.Cell(5, 1).Range.Text = '�����ˣ������˴����ˣ�����';
DTI.Cell(5, 3).Range.Text = '�Ա�';
DTI.Cell(5, 5).Range.Text = '����';
DTI.Cell(6, 1).Range.Text = '���֤��';
DTI.Cell(6, 3).Range.Text = '��λ��ְ��';
DTI.Cell(7, 1).Range.Text = 'סַ';
DTI.Cell(7, 3).Range.Text = '��ϵ�绰';
DTI.Cell(8, 1).Range.Text = '��������';
DTI.Cell(8, 3).Range.Text = '��λ��ְ��';
DTI.Cell(9, 1).Range.Text = '�� ¼ ��';
DTI.Cell(9, 3).Range.Text = '��λ��ְ��';


  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  
  text_temp='���飨��飩����������';
  set(selection, 'Text',text_temp);
  selection.Font.Size=15;%�����ı������С
  selection.Font.Bold=0;%�����ı�����
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  
  text_temp=['      ��   ��  ��   ʱ   �֣���ʻԱ '  ' �ڳ��������й�·�����·��Ա '...
            ' �� ''����S311ʡ��K234+650������          ����ʻ�ĳ��ƺ�Ϊ ' ...
            '�Ļ����ĳ������أ��������ߣ����������˼�⣬����⣬�ó���������       �֡� '...
      '___________________________________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'...
      '___________________________________________________________________'];
  set(selection, 'Text',text_temp);
  selection.Font.Size=13;%�����ı������С
  selection.Font.Bold=4;%�����ı�����
  selection.Font.underline = 2;
  set(paragraphformat, 'LineSpacing',23);
  
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  text_temp='�����˻��������ǩ����_____________���飨��飩��ǩ����_____________';
  set(selection, 'Text',text_temp);
  selection.Font.Size=15;%�����ı������С
  selection.Font.Bold=0;%�����ı�����
  set(paragraphformat, 'LineSpacing',30);
  
  end_of_doc = get(content,'end');%����ƶ������
  set(selection,'Start',end_of_doc); 
  selection.MoveDown;
  selection.TypeParagraph;
  text_temp='��������ǩ����_____________________��¼��ǩ����____________________';
  set(selection, 'Text',text_temp);

  
  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  end_of_doc = get(content,'end');%����ƶ������
  set(selection,'Start',end_of_doc); 
  text_temp='��ע��';
  set(selection, 'Text',text_temp);
  
  end_of_doc = get(content,'end');%����ƶ������
  set(selection,'Start',end_of_doc); 
    selection.MoveDown;
  selection.TypeParagraph;
  
    %����ҳ��ʼ
  
    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  
  text_temp='��  ��  ��  ��  ͼ';
  set(selection, 'Text',text_temp);


  selection.Font.Size=16;%�����ı������С
  selection.Font.Bold=4;%�����ı�����
  set(paragraphformat, 'LineSpacing',23);
  set(paragraphformat, 'Alignment','wdAlignParagraphCenter');

  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  
  text_temp='���ɣ� ';
  set(selection, 'Text',text_temp);
  selection.Font.Size=15;%�����ı������С
  selection.Font.Bold=0;%�����ı�����
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  
  text_temp='_______________________�������Գ���������ʻ��·��__________________ ';
  set(selection, 'Text',text_temp);
  selection.Font.Size=13;%�����ı������С
  selection.Font.Bold=4;%�����ı�����
  selection.Font.underline = 2;
  
  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  
    %��ʼ������ҳ��ͼ
  selection.Font.Bold=0;%�����ı�����
  selection.Font.underline = 0;
  selection.Font.Size=15;%�����ı������С
  Tables=document.Tables.Add(selection.Range,3,1);
 
%���ñ߿�
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

  DTI.Cell(2, 1).Range.Text = '�����˻������ǩ��:';
  DTI.Cell(3, 1).Range.Text = '����(���)��ǩ��:';

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  text_temp='��ע�� ';
  set(selection, 'Text',text_temp);
  selection.Font.Size=15;%�����ı������С
  selection.Font.Bold=0;%�����ı�����
  
    end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  
  %% �ڰ�ҳ��ʼ
  
   end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  
  text_temp='ճ   ��   ר   ҳ';
  set(selection, 'Text',text_temp);


  selection.Font.Size=16;%�����ı������С
  selection.Font.Bold=4;%�����ı�����
  set(paragraphformat, 'LineSpacing',23);
  set(paragraphformat, 'Alignment','wdAlignParagraphCenter');

  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  
  %��ʼ���ڰ�ҳ��ͼ
  selection.Font.Bold=0;%�����ı�����
  selection.Font.underline = 0;
  selection.Font.Size=15;%�����ı������С
  Tables=document.Tables.Add(selection.Range,1,1);
 
%���ñ߿�
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
  DTI.Cell(1, 1).Range.Text = '��  ��  ��  ��';

  end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  set(paragraphformat, 'Alignment','wdAlignParagraphLeft');
  
  selection.MoveDown;
  selection.TypeParagraph;
  
 %% �ھ�ҳ��ʼ
  
   end_of_doc = get(content,'end');
  set(selection,'Start',end_of_doc);
  
  text_temp='ճ   ��   ר   ҳ';
  set(selection, 'Text',text_temp);


  selection.Font.Size=16;%�����ı������С
  selection.Font.Bold=4;%�����ı�����
  set(paragraphformat, 'LineSpacing',23);
  set(paragraphformat, 'Alignment','wdAlignParagraphCenter');

  selection.MoveDown;
  selection.TypeParagraph;
  selection.TypeParagraph;
  
  %��ʼ���ڰ�ҳ��ͼ
  selection.Font.Bold=0;%�����ı�����
  selection.Font.underline = 0;
  selection.Font.Size=15;%�����ı������С
  Tables=document.Tables.Add(selection.Range,2,1);
 
%���ñ߿�
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

  DTI.Cell(1, 1).Range.Text = '��ʻԱ���֤����ӡ��';
  DTI.Cell(2, 1).Range.Text = '�������֤����ӡ��';

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

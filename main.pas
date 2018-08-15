// (c) 2018 Alexander Pchelintsev, pchelintsev.an@yandex.ru

unit main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, ExtCtrls, ComObj, Menus;

type
  TFormInput = class(TForm)
    LabeledEdit1: TLabeledEdit;
    Label1: TLabel;
    DateTimePicker1: TDateTimePicker;
    Label2: TLabel;
    DateTimePicker2: TDateTimePicker;
    LabeledEdit2: TLabeledEdit;
    Label3: TLabel;
    ComboBox1: TComboBox;
    LabeledEdit3: TLabeledEdit;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    StatusBar1: TStatusBar;
    LabeledEdit4: TLabeledEdit;
    LabeledEdit5: TLabeledEdit;
    LabeledEdit6: TLabeledEdit;
    Edit1: TEdit;
    CheckBox1: TCheckBox;
    LabeledEdit7: TLabeledEdit;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    Label4: TLabel;
    Button4: TButton;
    OpenDialog1: TOpenDialog;
    LabeledEdit8: TLabeledEdit;
    LabeledEdit9: TLabeledEdit;
    CheckBox2: TCheckBox;
    LabeledEdit10: TLabeledEdit;
    LabeledEdit11: TLabeledEdit;
    Button5: TButton;
    Button6: TButton;
    Edit2: TEdit;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure N4Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
  private
    { Private declarations }
    procedure CreteListStudents(FN: String);
    procedure StringReplace(W: OLEVariant; SearchStr,ReplaceStr: String;
                            ReplaceAll: Boolean);
    procedure SaveToFile;
    procedure ClickToButton(B: TButton);
    function InvString(fio: String): String;
  public
    { Public declarations }
  end;

  TInfoStudent = record
    Fam   : String;
    IO    : String;
    Theme : String;
    ocenka: String;
    page  : String
  end;

var
  FormInput: TFormInput;

implementation

{$R *.dfm}

var
  studs: array of TInfoStudent;
  n    : Integer = 0;

  fname_IUL1      : String = '';
  fname_IUL2      : String = '';
  fname_kr1       : String = '';
  fname_kr2       : String = '';
  fname_Perechen  : String = '';
  fname_stick     : String = '';
  fname_ved       : String = '';
  ex_file_IUL1    : Boolean = false;
  ex_file_IUL2    : Boolean = false;
  ex_file_kr1     : Boolean = false;
  ex_file_kr2     : Boolean = false;
  ex_file_Perechen: Boolean = false;
  ex_file_stick   : Boolean = false;
  ex_file_ved     : Boolean = false;
  path            : String = '';
  config_path     : String = '';

procedure TFormInput.CreteListStudents(FN: String);
var
  XL       : OLEVariant;
  XLRun    : Boolean;
  s        : AnsiString;
  ist,i,j,k: Integer;
  Name     : AnsiString;
  buf      : AnsiString;
  status   : AnsiString;
  flag     : Boolean;
begin
  studs:=nil;
  n:=0;
  Button1.Enabled:=false;
  Button2.Enabled:=false;
  Button3.Enabled:=false;
  Button4.Enabled:=false;
  Button5.Enabled:=false;
  Button6.Enabled:=false;
  XLRun:=false;
  try
    StatusBar1.SimpleText:='Запуск MS Excel, формирование списка студентов. Ждите...';
    XL:=CreateOleObject('Excel.Application');
    XLRun:=true;
    XL.WorkBooks.Open(FN);
    ist:=1;
    while true do
    begin
      s:=AnsiString(XL.WorkBooks[1].WorkSheets[1].Cells[ist,1].Value);
      if s = '' then
        break;
      inc(n);
      SetLength(studs, n);
      status:=AnsiString(XL.WorkBooks[1].WorkSheets[1].Cells[ist,2].Value);
      Name:='';
      if status <> '*' then
      begin
        flag:=false;
        for i:=2 to Length(s) do
          if s[i] in ['А'..'Я'] then
          begin
            flag:=true;
            break
          end;
        if flag then
          Name:=s[i] + '.';
        flag:=false;
        for j:=i+1 to Length(s) do
          if s[j] in ['А'..'Я'] then
          begin
            flag:=true;
            break
          end;
        if flag then
          Name:=Name + s[j] + '.';
        flag:=false;
        for k:=j+1 to Length(s) do
          if s[k] in ['А'..'Я'] then
          begin
            flag:=true;
            break
          end;
        if flag then
          Name:=Name + s[k] + '.';
        for j:=i downto 1 do
          if s[j] in ['а'..'я'] then
            break;
        buf:=Copy(s, 1, j);
        s:=buf
      end;
      studs[n-1].Fam   :=s;
      studs[n-1].IO    :=Name;
      studs[n-1].Theme :=AnsiString(XL.WorkBooks[1].WorkSheets[1].Cells[ist,3].Value);
      studs[n-1].ocenka:=AnsiString(XL.WorkBooks[1].WorkSheets[1].Cells[ist,4].Value);
      studs[n-1].page  :=AnsiString(XL.WorkBooks[1].WorkSheets[1].Cells[ist,5].Value);
      inc(ist)
    end;
    XL.Quit;
    XL:=Unassigned;
    Button1.Enabled:=ex_file_IUL1 or ex_file_IUL2;
    Button2.Enabled:=ex_file_Perechen;
    Button4.Enabled:=ex_file_stick;
    Button3.Enabled:=ex_file_ved;
    Button6.Enabled:=ex_file_kr1 or ex_file_kr2;
    Button5.Enabled:=ex_file_IUL1 or ex_file_IUL2 or ex_file_Perechen or
                     ex_file_stick or ex_file_ved or ex_file_kr1 or ex_file_kr2;
    StatusBar1.SimpleText:='Готово'
  except
    MessageDlg('Ошибка работы с xlsx-файлом',mtError,[mbOk],0);
    try
      if XLRun then XL.Quit
    except
    end;
    XL:=Unassigned;
    StatusBar1.SimpleText:=''
  end
end;

procedure TFormInput.StringReplace(W: OLEVariant; SearchStr,ReplaceStr: String;
                                   ReplaceAll: Boolean);
begin
  try
    W.Selection.Find.ClearFormatting;
    W.Selection.Find.Text:=SearchStr;
    W.Selection.Find.Replacement.Text:=ReplaceStr;
    W.Selection.Find.Forward:=true;
    W.Selection.Find.Wrap:=1;
    W.Selection.Find.Format:=false;
    W.Selection.Find.MatchCase:=false;
    W.Selection.Find.MatchWholeWord:=false;
    W.Selection.Find.MatchWildcards:=false;
    W.Selection.Find.MatchSoundsLike:=false;
    W.Selection.Find.MatchAllWordForms:=false;
    if ReplaceAll then W.Selection.Find.Execute(Replace:=2)
    else W.Selection.Find.Execute(Replace:=1)
  except
  end
end;

procedure TFormInput.Button1Click(Sender: TObject);
var
  WRun     : Boolean;
  W        : OLEVariant;
  i        : Integer;
  numstud  : String;
  namefile : String;
  FIOstud  : String;
  fname_IUL: String;
  names_fl : array of String;
begin
  WRun:=false;
  try
    StatusBar1.SimpleText:='Запуск MS Word, формирование ИУЛ. Ждите...';
    W:=CreateOleObject('Word.Application');
    WRun:=true;

    if CheckBox2.Checked then
      if ex_file_IUL2 then
        fname_IUL:=fname_IUL2
      else fname_IUL:=fname_IUL1
    else
      if ex_file_IUL1 then
        fname_IUL:=fname_IUL1
      else fname_IUL:=fname_IUL2;

    SetLength(names_fl, n);
    for i:=1 to n do
    begin
      W.Documents.Open(fname_IUL);
      StringReplace(W, '{napr}', LabeledEdit4.Text, true);
      numstud:=IntToStr(i);
      if Length(numstud) = 1 then
        numstud:='0'+numstud;
      StringReplace(W, '{numstud}', numstud, true);
      StringReplace(W, '{theme}', studs[i-1].Theme, true);
      StringReplace(W, '{enddate}', DateToStr(DateTimePicker2.DateTime), true);
      StringReplace(W, '{begindt}', DateToStr(DateTimePicker1.DateTime), true);
      if studs[i-1].IO = '' then
        FIOstud:=studs[i-1].Fam
      else
        FIOstud:=studs[i-1].Fam + ' ' + studs[i-1].IO;
      StringReplace(W, '{fiostud}', FIOstud, true);
      StringReplace(W, '{ocenka}', studs[i-1].ocenka, true);
      StringReplace(W, '{ng}', LabeledEdit7.Text, true);
      StringReplace(W, '{fioruk}', LabeledEdit3.Text, true);
      StringReplace(W, '{group}', LabeledEdit1.Text, true);
      if Edit2.Text<>'' then
        StringReplace(W, '{kafname}', Edit2.Text, true)
      else
        StringReplace(W, '{kafname}', LabeledEdit8.Text, true);
      StringReplace(W, '{zavkaf}', LabeledEdit9.Text, true);
      if CheckBox2.Checked then
      begin
        StringReplace(W, '{fiokomis2}', LabeledEdit10.Text, true);
        StringReplace(W, '{fiokomis3}', LabeledEdit11.Text, true)
      end;
      namefile:=path + 'ТГТУ ' + LabeledEdit4.Text + '.' + LabeledEdit7.Text +
                numstud + ' КР УЛ ' + FIOstud + '.docx';
      names_fl[i-1]:=namefile;
      W.ActiveDocument.SaveAs(FileName:=namefile, FileFormat:=16);
      W.ActiveDocument.Close
    end;
    if CheckBox1.Checked then
    begin
      W.Documents.Open(names_fl[0]);
      for i:=2 to n do
      begin
        W.Selection.EndKey(Unit:=6);
        W.Selection.InsertBreak(Type:=7);
        W.Selection.InsertFile(FileName:=names_fl[i-1]);
      end;
      W.ActiveDocument.SaveAs(FileName:=path + 'ТГТУ ' + LabeledEdit4.Text + ' ' +
                                        LabeledEdit1.Text + ' КР УЛ.docx',
                              FileFormat:=16);
      W.ActiveDocument.Close;
      for i:=1 to n do
        DeleteFile(names_fl[i-1])
    end;
    W.Quit;
    W:=Unassigned;
    StatusBar1.SimpleText:='Готово. Результирующие файлы находятся в папке docs'
  except
    MessageDlg('Ошибка при работе с MS Word',mtError,[mbOk],0);
    try
      if WRun then W.Quit
    except
    end;
    W:=Unassigned;
    StatusBar1.SimpleText:=''
  end;
  names_fl:=nil
end;

procedure TFormInput.Button2Click(Sender: TObject);
var
  WRun              : Boolean;
  W                 : OLEVariant;
  i                 : Integer;
  numstud           : String;
  namefile          : String;
  FIOstud1, FIOstud2: String;
  names_fl          : array of String;
begin
  WRun:=false;
  try
    StatusBar1.SimpleText:='Запуск MS Word, формирование документов. Ждите...';
    W:=CreateOleObject('Word.Application');
    WRun:=true;
    SetLength(names_fl, n);
    for i:=1 to n do
    begin
      W.Documents.Open(fname_Perechen);
      StringReplace(W, '{napr}', LabeledEdit4.Text, true);
      numstud:=IntToStr(i);
      if Length(numstud) = 1 then
        numstud:='0'+numstud;
      StringReplace(W, '{numstud}', numstud, true);
      if studs[i-1].IO = '' then
      begin
        FIOstud1:=studs[i-1].Fam;
        FIOstud2:=studs[i-1].Fam
      end
      else
      begin
        FIOstud1:=studs[i-1].Fam + ' ' + studs[i-1].IO;
        FIOstud2:=studs[i-1].IO + ' ' + studs[i-1].Fam
      end;
      StringReplace(W, '{fiostud}', FIOstud2, true);
      StringReplace(W, '{page}', studs[i-1].page, true);
      StringReplace(W, '{ng}', LabeledEdit7.Text, true);
      StringReplace(W, '{fioarch}', InvString(LabeledEdit2.Text), true);
      namefile:=path + 'ТГТУ ' + LabeledEdit4.Text + '.' + LabeledEdit7.Text + numstud +
                ' КР Перечень ' + FIOstud1 + '.docx';
      names_fl[i-1]:=namefile;
      W.ActiveDocument.SaveAs(FileName:=namefile, FileFormat:=16);
      W.ActiveDocument.Close
    end;
    if CheckBox1.Checked then
    begin
      W.Documents.Open(names_fl[0]);
      for i:=2 to n do
      begin
        W.Selection.EndKey(Unit:=6);
        W.Selection.InsertBreak(Type:=7);
        W.Selection.InsertFile(FileName:=names_fl[i-1]);
      end;
      W.ActiveDocument.SaveAs(FileName:=path + 'ТГТУ ' + LabeledEdit4.Text + ' ' +
                                        LabeledEdit1.Text + ' КР Перечень.docx',
                              FileFormat:=16);
      W.ActiveDocument.Close;
      for i:=1 to n do
        DeleteFile(names_fl[i-1])
    end;
    W.Quit;
    W:=Unassigned;
    StatusBar1.SimpleText:='Готово. Результирующие файлы находятся в папке docs'
  except
    MessageDlg('Ошибка при работе с MS Word',mtError,[mbOk],0);
    try
      if WRun then W.Quit
    except
    end;
    W:=Unassigned;
    StatusBar1.SimpleText:=''
  end;
  names_fl:=nil
end;

procedure TFormInput.SaveToFile;
var
  f: TextFile;
begin
  studs:=nil;
  AssignFile(f, config_path);
  {$I-} Rewrite(f); {$I+}
  if IOResult = 0 then
  begin
    WriteLn(f, LabeledEdit1.Text);
    WriteLn(f, LabeledEdit7.Text);
    WriteLn(f, LabeledEdit2.Text);
    WriteLn(f, ComboBox1.Text);
    WriteLn(f, LabeledEdit3.Text);
    WriteLn(f, LabeledEdit4.Text);
    WriteLn(f, Edit1.Text);
    WriteLn(f, LabeledEdit5.Text);
    WriteLn(f, LabeledEdit6.Text);
    WriteLn(f, DateToStr(DateTimePicker1.DateTime));
    WriteLn(f, DateToStr(DateTimePicker2.DateTime));
    WriteLn(f, LabeledEdit8.Text);
    WriteLn(f, Edit2.Text);
    WriteLn(f, LabeledEdit9.Text);
    if CheckBox1.Checked then
      WriteLn(f, '1')
    else
      WriteLn(f, '0');
    if CheckBox2.Checked then
    begin
      WriteLn(f, '1');
      WriteLn(f, LabeledEdit10.Text);
      WriteLn(f, LabeledEdit11.Text)
    end
    else
      WriteLn(f, '0');
    CloseFile(f)
  end
end;

procedure TFormInput.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  SaveToFile
end;

procedure TFormInput.N4Click(Sender: TObject);
begin
  SaveToFile;
  Application.Terminate
end;

procedure TFormInput.FormCreate(Sender: TObject);
var
  f: TextFile;
  s: String;
begin
  path:=ExtractFilePath(Application.ExeName);

  config_path    := path + 'config.txt';
  fname_IUL1     := path + 'templates\IUL1.docx';
  fname_IUL2     := path + 'templates\IUL2.docx';
  fname_kr1      := path + 'templates\kr1.docx';
  fname_kr2      := path + 'templates\kr2.docx';
  fname_Perechen := path + 'templates\archive.docx';
  fname_stick    := path + 'templates\sticker.docx';
  fname_ved      := path + 'templates\vedomost.docx';

  ex_file_IUL1     := FileExists(fname_IUL1);
  ex_file_IUL2     := FileExists(fname_IUL2);
  ex_file_kr1      := FileExists(fname_kr1);
  ex_file_kr2      := FileExists(fname_kr2);
  ex_file_Perechen := FileExists(fname_Perechen);
  ex_file_stick    := FileExists(fname_stick);
  ex_file_ved      := FileExists(fname_ved);

  path:=path + 'docs\';

  if not DirectoryExists(path) then
    CreateDir(path);

  AssignFile(f, config_path);
  {$I-} Reset(f); {$I+}
  if IOResult = 0 then
  begin
    ReadLn(f, s);
    LabeledEdit1.Text:=s;
    ReadLn(f, s);
    LabeledEdit7.Text:=s;
    ReadLn(f, s);
    LabeledEdit2.Text:=s;
    ReadLn(f, s);
    ComboBox1.Text:=s;
    ReadLn(f, s);
    LabeledEdit3.Text:=s;
    ReadLn(f, s);
    LabeledEdit4.Text:=s;
    ReadLn(f, s);
    Edit1.Text:=s;
    ReadLn(f, s);
    LabeledEdit5.Text:=s;
    ReadLn(f, s);
    LabeledEdit6.Text:=s;
    try
      ReadLn(f, s);
      DateTimePicker1.DateTime:=StrToDate(s);
      ReadLn(f, s);
      DateTimePicker2.DateTime:=StrToDate(s);
    except
    end;
    ReadLn(f, s);
    LabeledEdit8.Text:=s;
    ReadLn(f, s);
    Edit2.Text:=s;
    ReadLn(f, s);
    LabeledEdit9.Text:=s;
    ReadLn(f, s);
    CheckBox1.Checked:=s='1';
    ReadLn(f, s);
    CheckBox2.Checked:=s='1';
    if CheckBox2.Checked then
    begin
      ReadLn(f, s);
      LabeledEdit10.Text:=s;
      ReadLn(f, s);
      LabeledEdit11.Text:=s
    end;
    CloseFile(f)
  end
end;

procedure TFormInput.N2Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
    CreteListStudents(OpenDialog1.FileName)
end;

procedure TFormInput.Button4Click(Sender: TObject);
var
  WRun    : Boolean;
  W       : OLEVariant;
  i       : Integer;
  numstud : String;
  namefile: String;
  FIOstud : String;
  names_fl: array of String;
begin
  WRun:=false;
  try
    StatusBar1.SimpleText:='Запуск MS Word, формирование этикеток. Ждите...';
    W:=CreateOleObject('Word.Application');
    WRun:=true;
    SetLength(names_fl, n);
    for i:=1 to n do
    begin
      W.Documents.Open(fname_stick);
      StringReplace(W, '{napr}', LabeledEdit4.Text, true);
      numstud:=IntToStr(i);
      if Length(numstud) = 1 then
        numstud:='0'+numstud;
      StringReplace(W, '{numstud}', numstud, true);
      if studs[i-1].IO = '' then
        FIOstud:=studs[i-1].Fam
      else
        FIOstud:=studs[i-1].Fam + ' ' + studs[i-1].IO;
      StringReplace(W, '{fiostud}', FIOstud, true);
      StringReplace(W, '{ng}', LabeledEdit7.Text, true);
      StringReplace(W, '{group}', LabeledEdit1.Text, true);
      StringReplace(W, '{theme}', studs[i-1].Theme, true);
      StringReplace(W, '{year}',FormatDateTime('yyyy',DateTimePicker2.DateTime),
                    true);
      namefile:=path + 'ТГТУ ' + LabeledEdit4.Text + '.' + LabeledEdit7.Text + numstud +
                ' КР Этикетка ' + FIOstud + '.docx';
      names_fl[i-1]:=namefile;
      W.ActiveDocument.SaveAs(FileName:=namefile, FileFormat:=16);
      W.ActiveDocument.Close
    end;
    if CheckBox1.Checked then
    begin
      W.Documents.Open(names_fl[0]);
      for i:=2 to n do
      begin
        W.Selection.EndKey(Unit:=6);
        W.Selection.InsertBreak(Type:=7);
        W.Selection.InsertFile(FileName:=names_fl[i-1]);
      end;
      W.ActiveDocument.SaveAs(FileName:=path + 'ТГТУ ' + LabeledEdit4.Text + ' ' +
                                        LabeledEdit1.Text + ' КР Этикетки.docx',
                              FileFormat:=16);
      W.ActiveDocument.Close;
      for i:=1 to n do
        DeleteFile(names_fl[i-1])
    end;
    W.Quit;
    W:=Unassigned;
    StatusBar1.SimpleText:='Готово. Результирующие файлы находятся в папке docs'
  except
    MessageDlg('Ошибка при работе с MS Word',mtError,[mbOk],0);
    try
      if WRun then W.Quit
    except
    end;
    W:=Unassigned;
    StatusBar1.SimpleText:=''
  end;
  names_fl:=nil
end;

procedure TFormInput.Button3Click(Sender: TObject);
var
  WRun   : Boolean;
  W      : OLEVariant;
  i      : Integer;
  FIOstud: String;
begin
  WRun:=false;
  try
    StatusBar1.SimpleText:='Запуск MS Word, формирование ведомости. Ждите...';
    W:=CreateOleObject('Word.Application');
    WRun:=true;
    W.Documents.Open(fname_ved);
    StringReplace(W, '{napr}', LabeledEdit4.Text, true);
    StringReplace(W, '{naprname}', Edit1.Text, true);
    StringReplace(W, '{institute}', ComboBox1.Text, true);
    StringReplace(W, '{sem}', LabeledEdit6.Text, true);
    StringReplace(W, '{disciplina}', LabeledEdit5.Text, true);
    StringReplace(W, '{group}', LabeledEdit1.Text, true);
    StringReplace(W, '{fioruk}', InvString(LabeledEdit3.Text), true);
    StringReplace(W, '{zavkaf}', InvString(LabeledEdit9.Text), true);
    if n>=1 then
    begin
      if studs[0].IO = '' then
        FIOstud:=studs[0].Fam
      else
        FIOstud:=studs[0].Fam + ' ' + studs[0].IO;
      StringReplace(W, '{nach}', FIOstud, false);
      W.Selection.MoveRight(Unit:=$C);
      W.Selection.TypeText(Text:=studs[0].Theme);
      for i:=2 to n do
      begin
        W.Selection.MoveRight(Unit:=$C);
        W.Selection.MoveRight(Unit:=$C);
        if studs[i-1].IO = '' then
          FIOstud:=studs[i-1].Fam
        else
          FIOstud:=studs[i-1].Fam + ' ' + studs[i-1].IO;
        W.Selection.TypeText(Text:=FIOstud);
        W.Selection.MoveRight(Unit:=$C);
        W.Selection.TypeText(Text:=studs[i-1].Theme)
      end
    end
    else
      StringReplace(W, '{nach}', '', false);
    W.ActiveDocument.SaveAs(FileName:=path + 'ТГТУ ' + LabeledEdit4.Text + ' ' +
                                      LabeledEdit1.Text + ' КР Ведомость.docx',
                            FileFormat:=16);
    W.ActiveDocument.Close;
    W.Quit;
    W:=Unassigned;
    StatusBar1.SimpleText:='Готово. Результирующий файл находится в папке docs'
  except
    MessageDlg('Ошибка при работе с MS Word',mtError,[mbOk],0);
    try
      if WRun then W.Quit
    except
    end;
    W:=Unassigned;
    StatusBar1.SimpleText:=''
  end
end;

procedure TFormInput.CheckBox2Click(Sender: TObject);
begin
  if CheckBox2.Checked then
  begin
    LabeledEdit10.Enabled:=true;
    if LabeledEdit10.Text='' then
      LabeledEdit10.Text:=LabeledEdit9.Text;
    LabeledEdit11.Enabled:=true
  end
  else
  begin
    LabeledEdit10.Enabled:=false;
    LabeledEdit11.Enabled:=false
  end
end;

function TFormInput.InvString(fio: String): String;
var
  res : String;
  i,j : Integer;
  flag: Boolean;
begin
  flag:=false;
  for i:=2 to Length(fio) do
    if fio[i]=' ' then
    begin
      flag:=true;
      break
    end;
  if flag then
  begin
    res:=Copy(fio, 1, i-1);
    flag:=false;
    for j:=i to Length(fio) do
      if fio[j]<>' ' then
      begin
        flag:=true;
        break
      end;
    if flag then
      res:=Copy(fio, j, Length(fio)-j+1) + ' ' + res
  end
  else
    res:=fio;
  InvString:=res
end;

procedure TFormInput.Button6Click(Sender: TObject);

var
  WRun    : Boolean;
  W       : OLEVariant;
  i       : Integer;
  numstud : String;
  namefile: String;
  fname_kr: String;
  FIOstud1, FIOstud2: String;
begin
  WRun:=false;
  try
    StatusBar1.SimpleText:='Запуск MS Word, формирование шаблонов работ. Ждите...';
    W:=CreateOleObject('Word.Application');
    WRun:=true;

    if CheckBox2.Checked then
      if ex_file_kr2 then
        fname_kr:=fname_kr2
      else fname_kr:=fname_kr1
    else
      if ex_file_kr1 then
        fname_kr:=fname_kr1
      else fname_kr:=fname_kr2;

    for i:=1 to n do
    begin
      W.Documents.Open(fname_kr);

      if LabeledEdit8.Text='' then
        StringReplace(W, '{kafname}', Edit2.Text, true)
      else
        StringReplace(W, '{kafname}', LabeledEdit8.Text, true);
      StringReplace(W, '{zavkaf}', InvString(LabeledEdit9.Text), true);

      StringReplace(W, '{dend}',FormatDateTime('dd',DateTimePicker2.DateTime),
                    true);
      StringReplace(W, '{mend}',FormatDateTime('mm',DateTimePicker2.DateTime),
                    true);
      StringReplace(W, '{yend}',FormatDateTime('yyyy',DateTimePicker2.DateTime),
                    true);

      StringReplace(W, '{disciplina}', LabeledEdit5.Text, true);
      StringReplace(W, '{theme}', studs[i-1].Theme, true);
      StringReplace(W, '{napr}', LabeledEdit4.Text, true);
      StringReplace(W, '{naprname}', Edit1.Text, true);


      if studs[i-1].IO = '' then
      begin
        FIOstud1:=studs[i-1].Fam;
        FIOstud2:=studs[i-1].Fam
      end
      else
      begin
        FIOstud1:=studs[i-1].Fam + ' ' + studs[i-1].IO;
        FIOstud2:=studs[i-1].IO + ' ' + studs[i-1].Fam
      end;
      StringReplace(W, '{fiostud}', FIOstud2, true);

      StringReplace(W, '{group}', LabeledEdit1.Text, true);
      StringReplace(W, '{ng}', LabeledEdit7.Text, true);

      numstud:=IntToStr(i);
      if Length(numstud) = 1 then
        numstud:='0'+numstud;
      StringReplace(W, '{numstud}', numstud, true);

      StringReplace(W, '{fioruk}', InvString(LabeledEdit3.Text), true);
      StringReplace(W, '{ocenka}', studs[i-1].ocenka, true);

      if CheckBox2.Checked then
      begin
        StringReplace(W, '{fiokomis2}', InvString(LabeledEdit10.Text), true);
        StringReplace(W, '{fiokomis3}', InvString(LabeledEdit11.Text), true)
      end;

      StringReplace(W, '{year}',FormatDateTime('yyyy',DateTimePicker2.DateTime),
                    true);

      StringReplace(W, '{dbeg}',FormatDateTime('dd',DateTimePicker1.DateTime),
                    true);
      StringReplace(W, '{mbeg}',FormatDateTime('mm',DateTimePicker1.DateTime),
                    true);
      StringReplace(W, '{ybeg}',FormatDateTime('yyyy',DateTimePicker1.DateTime),
                    true);

      namefile:=path + 'ТГТУ ' + LabeledEdit4.Text + '.' + LabeledEdit7.Text +
                numstud + ' КР ДЭ ' + FIOstud1 + '.docx';;

      W.ActiveDocument.SaveAs(FileName:=namefile, FileFormat:=16);
      W.ActiveDocument.Close
    end;
    W.Quit;
    W:=Unassigned;
    StatusBar1.SimpleText:='Готово. Результирующие файлы находятся в папке docs'
  except
    MessageDlg('Ошибка при работе с MS Word',mtError,[mbOk],0);
    try
      if WRun then W.Quit
    except
    end;
    W:=Unassigned;
    StatusBar1.SimpleText:=''
  end
end;

procedure TFormInput.ClickToButton(B: TButton);
begin
  if B.Enabled then
    B.Click
end;

procedure TFormInput.Button5Click(Sender: TObject);
begin
  ClickToButton(Button1);
  ClickToButton(Button2);
  ClickToButton(Button3);
  ClickToButton(Button4);
  ClickToButton(Button6)
end;

end.

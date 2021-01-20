unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, StdCtrls, Actions, ActnList, ImgList, ExtDlgs, ComCtrls,
  // TB2K
  TB2Item, TB2Toolbar, TB2Dock, TB2ExtItems,
  // SpTBXLib
  SpTBXSkins, SpTBXItem, SpTBXDkPanels, SpTBXTabs, SpTBXEditors, SpTBXControls,
  SpTBXExtEditors, SpTBXImageList;

type
  TForm1 = class(TForm)
    SpTBXDock1: TSpTBXDock;
    SpTBXMultiDock1: TSpTBXMultiDock;
    SpTBXMultiDock3: TSpTBXMultiDock;
    SpTBXDockablePanel1: TSpTBXDockablePanel;
    SpTBXToolbar1: TSpTBXToolbar;
    Panel1: TPanel;
    SpTBXTabControl1: TSpTBXTabControl;
    SpTBXRightAlignSpacerItem1: TSpTBXRightAlignSpacerItem;
    tabClose: TSpTBXItem;
    SpTBXSubmenuItem1: TSpTBXSubmenuItem;
    SpTBXSubmenuItem2: TSpTBXSubmenuItem;
    SpTBXItem2: TSpTBXItem;
    tabRight: TSpTBXItem;
    tabLeft: TSpTBXItem;
    SpTBXToolbar3: TSpTBXToolbar;
    SpTBXLabelItem3: TSpTBXLabelItem;
    SpTBXItem7: TSpTBXItem;
    subColor: TSpTBXSubmenuItem;
    SpTBXLabelItem4: TSpTBXLabelItem;
    subLang: TSpTBXSubmenuItem;
    subSkins: TSpTBXSubmenuItem;
    SpTBXSkinGroupItem1: TSpTBXSkinGroupItem;
    SpTBXTabControl2: TSpTBXTabControl;
    SpTBXTabItem1: TSpTBXTabItem;
    SpTBXTabSheet1: TSpTBXTabSheet;
    ActionList1: TActionList;
    Action1: TAction;
    Action2: TAction;
    SpTBXButton1: TSpTBXButton;
    SpTBXStatusBar1: TSpTBXStatusBar;
    SpTBXSeparatorItem2: TSpTBXSeparatorItem;
    hintLabel: TSpTBXLabelItem;
    SpTBXSeparatorItem4: TSpTBXSeparatorItem;
    subLang2: TSpTBXSubmenuItem;
    Image1: TImage;
    SpTBXSeparatorItem5: TSpTBXSeparatorItem;
    SpTBXSeparatorItem8: TSpTBXSeparatorItem;
    SpTBXTabItem3: TSpTBXTabItem;
    SpTBXTabSheet3: TSpTBXTabSheet;
    SpTBXGroupBox5: TSpTBXGroupBox;
    SpTBXLabel2: TSpTBXLabel;
    SpTBXLabel3: TSpTBXLabel;
    SpTBXLabel4: TSpTBXLabel;
    SpTBXLabel5: TSpTBXLabel;
    SpTBXToolbar2: TSpTBXToolbar;
    SpTBXItem1: TSpTBXItem;
    SpTBXSeparatorItem6: TSpTBXSeparatorItem;
    SpTBXItem3: TSpTBXItem;
    SpTBXItem4: TSpTBXItem;
    SpTBXSeparatorItem7: TSpTBXSeparatorItem;
    DP1: TSpTBXDockablePanel;
    DP2: TSpTBXDockablePanel;
    DP3: TSpTBXDockablePanel;
    SpTBXLabel1: TSpTBXLabel;
    SpTBXLabel8: TSpTBXLabel;
    SpTBXLabel9: TSpTBXLabel;
    SpTBXLabel10: TSpTBXLabel;
    SpTBXLabel11: TSpTBXLabel;
    SpTBXLabel12: TSpTBXLabel;
    SpTBXLabel13: TSpTBXLabel;
    SpTBXLabel15: TSpTBXLabel;
    SpTBXLabel16: TSpTBXLabel;
    SpTBXLabel17: TSpTBXLabel;
    Timer1: TTimer;
    SpTBXGroupBox6: TSpTBXGroupBox;
    progressDec: TSpTBXButton;
    progressInc: TSpTBXButton;
    progressAnimate: TSpTBXButton;
    SpTBXProgressBar1: TSpTBXProgressBar;
    progressFiles: TSpTBXCheckBox;
    SpTBXProgressBar2: TSpTBXProgressBar;
    TBControlItem2: TTBControlItem;
    SpTBXTrackBar1: TSpTBXTrackBar;
    SpTBXComboBox1: TSpTBXComboBox;
    TBControlItem3: TTBControlItem;
    SpTBXPanel2: TSpTBXPanel;
    Memo2: TMemo;
    SpTBXSplitter1: TSpTBXSplitter;
    SpTBXSplitter2: TSpTBXSplitter;
    SpTBXTabItem6: TSpTBXTabItem;
    SpTBXTabSheet6: TSpTBXTabSheet;
    SpTBXButtonEdit1: TSpTBXButtonEdit;
    SpTBXComboBox2: TSpTBXComboBox;
    SpTBXEdit1: TSpTBXEdit;
    SpTBXSpinEdit1: TSpTBXSpinEdit;
    LangListBox: TSpTBXListBox;
    SpTBXGroupBox1: TSpTBXGroupBox;
    SpTBXGroupBox2: TSpTBXGroupBox;
    trackTickmarks: TSpTBXRadioGroup;
    skinButton: TSpTBXSpeedButton;
    SpTBXDockablePanel2: TSpTBXDockablePanel;
    OpenDialog1: TOpenDialog;
    SpTBXSeparatorItem1: TSpTBXSeparatorItem;
    SpTBXColorPalette1: TSpTBXColorPalette;
    SpTBXItem5: TSpTBXItem;
    SpTBXSpinEditItem1: TSpTBXSpinEditItem;
    Panel2: TPanel;
    SpTBXSpeedButton1: TSpTBXSpeedButton;
    SpTBXSpeedButton2: TSpTBXSpeedButton;
    SpTBXSpeedButton3: TSpTBXSpeedButton;
    SpTBXColorEdit1: TSpTBXColorEdit;
    SpTBXFontComboBox1: TSpTBXFontComboBox;
    SpTBXGroupBox3: TSpTBXGroupBox;
    radiobuttonSkin1: TSpTBXRadioButton;
    radiobuttonSkin2: TSpTBXRadioButton;
    radiobuttonSkin3: TSpTBXRadioButton;
    procedure FormShow(Sender: TObject);
    procedure tabCloseClick(Sender: TObject);
    procedure ActionList1Update(Action: TBasicAction;
      var Handled: Boolean);
    procedure Action1Execute(Sender: TObject);
    procedure Action2Execute(Sender: TObject);
    procedure SpTBXTabControl2Resize(Sender: TObject);
    procedure rgSkinTypeClick(Sender: TObject);
    procedure hintLabelDrawHint(Sender: TObject;
      AHintBitmap: TBitmap; var AHint: string;
      var PaintDefault: Boolean);
    procedure progressDecClick(Sender: TObject);
    procedure progressIncClick(Sender: TObject);
    procedure progressAnimateClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure SpTBXProgressBar1ProgressChange(Sender: TObject;
      NewPosition: Integer);
    procedure progressFilesClick(Sender: TObject);
    procedure TrackbarTickMarksRadioClick(Sender: TObject);
    procedure skinButtonClick(Sender: TObject);
    procedure SpTBXSpeedButton1Draw(Sender: TObject; ACanvas: TCanvas;
      ARect: TRect; const PaintStage: TSpTBXPaintStage;
      var PaintDefault: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    FLastSkin: string;
    procedure LangClick(Sender: TObject);
    procedure WMSpSkinChange(var Message: TMessage); message WM_SPSKINCHANGE;
  public
    AppPath: string;
    IL: TSpTBXImageList;
  end;

var
  Form1: TForm1;

implementation

uses
  Themes, IOUtils;

{$R *.dfm}

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Utility functions }

function SpIDEBDSCommonDir(RADStudioIDENumber: Integer): string;
var
  S: string;
begin
  Result := '';
  // C:\Users\Public\Documents\Embarcadero\Studio\21.0
  S := IncludeTrailingPathDelimiter(TPath.GetSharedDocumentsPath) +
    'Embarcadero\Studio\' + IntToStr(RADStudioIDENumber) + '.0';
  if DirectoryExists(S) then
    Result := S;
end;

function SpGetDelphiStylesFolder: string;
var
  I: Integer;
  S: string;
begin
  Result := '';
  // XE2 = 9
  for I := 30 downto 9 do begin
    S := SpIDEBDSCommonDir(I) + '\Styles'; // C:\Users\Public\Documents\Embarcadero\Studio\21.0\Styles
    if DirectoryExists(S) then begin
      Result := S;
      Exit;
    end;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Form }

procedure TForm1.FormCreate(Sender: TObject);
begin
  FLastSkin := 'Aluminum';
  SkinManager.AddSkinNotification(Self);
  IL := TSpTBXImageList.Create(Self);
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  SkinManager.RemoveSkinNotification(Self);
  IL.Free;
end;

procedure TForm1.FormShow(Sender: TObject);
var
  A: TSpTBXItem;
  S: string;
  I: Integer;
begin
  // Add the Languages to the Languages menu item and TabControl
  if subLang.Count = 0 then
    for I := 0 to LangListBox.Items.Count - 1 do begin
      A := TSpTBXItem.Create(nil);
      try
        A.Caption := LangListBox.Items[I];
        A.GroupIndex := 100;
        A.AutoCheck := True;
        A.Tag := I;
        A.OnClick := LangClick;
        subLang.Add(A);
        with SpTBXTabControl1.Add(LangListBox.Items[I]) do
          Tag := I;
      except
        A.Free;
      end;
    end;

  // Select the first Language
  subLang.Items[0].Click;
  SpTBXTabControl1.ActiveTabIndex := 0;
  SpTBXTabControl2.ActiveTabIndex := 0;

  // Load default button Skin
  AppPath := IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName));

  // Load PNGs
  // Use ImageCollection/TVirtualImagelist on Rio and up
  IL.LoadGlyphs(AppPath + 'Glyphs');

  SpTBXTabControl1.Images := IL.ImageList;
  SpTBXToolbar3.Images := IL.ImageList;
  SpTBXStatusBar1.Images := IL.ImageList;
  DP2.Images := IL.ImageList;
  DP3.Images := IL.ImageList;

  // Initialize the link labels
  S := TPath.GetDocumentsPath;
  if Length(S) > 0 then begin
    SpTBXLabel4.LinkText := S;
    SpTBXLabel5.LinkText := 'explorer.exe';
    SpTBXLabel5.LinkTextParams := '/e, ' + S;
  end;

  // Init Skin Type
  radiobuttonSkin2.Enabled := CompilerVersion >= 23;  // Delphi XE2 or up
end;

procedure TForm1.LangClick(Sender: TObject);
var
  A: TSpTBXItem;
begin
  A := Sender as TSpTBXItem;
  subLang.Caption := A.Caption;
  subLang2.Caption := A.Caption;
  SpTBXTabControl1.ActiveTabIndex := subLang.IndexOf(A);
end;

procedure TForm1.tabCloseClick(Sender: TObject);
begin
  SpTBXTabControl1.Visible := False;
end;

procedure TForm1.ActionList1Update(Action: TBasicAction; var Handled: Boolean);
var
  L, R: Boolean;
begin
  SpTBXTabControl1.ScrollState(L, R);
  Action1.Enabled := L;
  Action2.Enabled := R;
end;

procedure TForm1.Action1Execute(Sender: TObject);
begin
  SpTBXTabControl1.ScrollLeft;
end;

procedure TForm1.Action2Execute(Sender: TObject);
begin
  SpTBXTabControl1.ScrollRight;
end;

procedure TForm1.WMSpSkinChange(var Message: TMessage);
begin
  if SkinManager.GetSkinType = sknSkin then
    radiobuttonSkin3.Checked := True
  else
    radiobuttonSkin1.Checked := True;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Options Panel }

procedure TForm1.rgSkinTypeClick(Sender: TObject);
var
  SkinType: TSpTBXSkinType;
begin
  SkinType := sknNone;
  if Sender = radiobuttonSkin1 then SkinType := sknWindows
  else if Sender = radiobuttonSkin2 then SkinType := sknDelphiStyle
  else if Sender = radiobuttonSkin3 then SkinType := sknSkin;

  if SkinType <> SkinManager.GetSkinType then
    case SkinType of
      sknWindows:
        begin
          if not SkinManager.IsDefaultSkin then begin
            FLastSkin := SkinManager.CurrentSkinName;
            SkinManager.SetSkin('Default');
          end;
        end;
      sknSkin:
        SkinManager.SetSkin(FLastSkin);
      sknDelphiStyle:
        begin
          OpenDialog1.InitialDir := SpGetDelphiStylesFolder;
          OpenDialog1.Filter := 'VCL Styles|*.vsf';
          if OpenDialog1.Execute then
            if FileExists(OpenDialog1.FileName) then begin
              SkinManager.SetToDefaultSkin;
              {$IF CompilerVersion >= 23} //for Delphi XE2 and up
              TStyleManager.SetStyle(TStyleManager.LoadFromFile(OpenDialog1.FileName));
              // Recreate the SkinGroupItem and add the selected style
              SpTBXSkinGroupItem1.Recreate;
              {$IFEND}
            end;
        end;
    end;
end;

procedure TForm1.skinButtonClick(Sender: TObject);
var
  S: string;
begin
  S := AppPath + 'Skins';
  if DirectoryExists(S) then
    OpenDialog1.InitialDir := S;
  OpenDialog1.Filter := 'Skin files|*.skn';

  if OpenDialog1.Execute then
    if FileExists(OpenDialog1.FileName) then begin
      // Load the skin file and add it to the SkinList
      S := SkinManager.AddSkinFromFile(OpenDialog1.FileName);
      if S <> '' then begin
        // Set the new skin
        FLastSkin := S;
        SkinManager.SetSkin(FLastSkin);
        // Recreate the SkinGroupItem and add the selected skin
        SpTBXSkinGroupItem1.Recreate;
        radiobuttonSkin3.Checked := True;
      end;
    end;
end;

procedure TForm1.SpTBXTabControl2Resize(Sender: TObject);
begin
  SpTBXButton1.Left := (SpTBXTabControl2.Width - SpTBXButton1.Width) div 2;
  SpTBXProgressBar1.Left := (SpTBXTabControl2.Width - SpTBXProgressBar1.Width) div 2;
  SpTBXTrackBar1.Left := (SpTBXTabControl2.Width - SpTBXTrackBar1.Width) div 2;
end;

procedure TForm1.SpTBXSpeedButton1Draw(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; const PaintStage: TSpTBXPaintStage; var PaintDefault: Boolean);
var
  SB: TSpTBXSpeedButton;
begin
  if PaintStage = pstPrePaint then begin
    PaintDefault := False;
    SB := TSpTBXSpeedButton(Sender);
    SpDrawXPHeader(ACanvas, ARect, SB.MouseInControl, SB.Pushed, CurrentPPI);
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ StatusBar }

procedure TForm1.hintLabelDrawHint(Sender: TObject;
  AHintBitmap: TBitmap; var AHint: string; var PaintDefault: Boolean);
var
  R, CaptionR, GlyphR: TRect;
  S: string;
  GSize: TSize;
  Margin: Integer;
begin
  PaintDefault := False;

  Margin := SpPPIScale(5, CurrentPPI);
  AHintBitmap.Canvas.Brush.Color := clInfoBk;
  AHintBitmap.Canvas.Font.Color := clInfoText;
  S := 'Selected Language: ' + subLang.Caption + #13#10 +
       'Selected Skin: ' + SkinManager.CurrentSkinName + #13#10 +
       'http://www.silverpointdevelopment.com';
  GSize := SpGetTextSize(AHintBitmap.Canvas.Handle, 'http://www.silverpointdevelopment.com', True);

  AHintBitmap.Width := Image1.Picture.Bitmap.Width + GSize.cx + Margin * 4;
  AHintBitmap.Height := Image1.Picture.Bitmap.Height + Margin * 2;
  R := Rect(0, 0, AHintBitmap.Width, AHintBitmap.Height);
  AHintBitmap.Canvas.FillRect(R);

  R := Rect(Margin, Margin, AHintBitmap.Width - Margin, AHintBitmap.Height - Margin);
  GSize.cx := Image1.Picture.Bitmap.Width;
  GSize.cy := Image1.Picture.Bitmap.Height;
  SpCalcXPText(AHintBitmap.Canvas, R, S, taLeftJustify, DT_WORDBREAK, GSize,
    ghlGlyphLeft, False, CurrentPPI, CaptionR, GlyphR);
  SpDrawXPText(AHintBitmap.Canvas, S, CaptionR, DT_WORDBREAK);
  AHintBitmap.Canvas.Draw(GlyphR.Left, GlyphR.Top, Image1.Picture.Bitmap);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ ProgressBar }

procedure TForm1.progressDecClick(Sender: TObject);
begin
  SpTBXProgressBar1.StepIt(-10);
  SpTBXProgressBar2.StepIt(-10);
end;

procedure TForm1.progressIncClick(Sender: TObject);
begin
  SpTBXProgressBar1.StepIt(10);
  SpTBXProgressBar2.StepIt(10);
end;

procedure TForm1.progressAnimateClick(Sender: TObject);
begin
  SpTBXProgressBar1.Position := 0;
  SpTBXProgressBar2.Position := 0;
  Timer1.Enabled := True;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
  SpTBXProgressBar1.StepIt;
  SpTBXProgressBar2.StepIt;
  if SpTBXProgressBar1.Position >= SpTBXProgressBar1.Max then
    Timer1.Enabled := False;
end;

procedure TForm1.SpTBXProgressBar1ProgressChange(Sender: TObject;
  NewPosition: Integer);
var
  I: Integer;
  S: string;
begin
  if progressFiles.Checked then begin
    I := (NewPosition div 10) - 1;
    if I < 0  then I := 0;
    if I > LangListBox.Items.Count - 1 then I := LangListBox.Items.Count - 1;
    S := 'C:\Lang\' + LangListBox.Items[I] + '.txt';
    SpTBXProgressBar1.Caption := S;
    SpTBXProgressBar2.Caption := S;
  end;
end;

procedure TForm1.ProgressFilesClick(Sender: TObject);
begin
  if progressFiles.Checked then
    SpTBXProgressBar1.CaptionType := pctDefault
  else
    SpTBXProgressBar1.CaptionType := pctPercentage;
  SpTBXProgressBar2.CaptionType := SpTBXProgressBar1.CaptionType;
  SpTBXProgressBar1ProgressChange(SpTBXProgressBar1, SpTBXProgressBar1.Position);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TrackBar }

procedure TForm1.TrackbarTickMarksRadioClick(Sender: TObject);
begin
  SpTBXTrackBar1.TickMarks := TSpTBXTickMark(trackTickmarks.ItemIndex);
end;

end.

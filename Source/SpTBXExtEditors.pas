unit SpTBXExtEditors;

{==============================================================================
Version 2.5.9

The contents of this file are subject to the SpTBXLib License; you may
not use or distribute this file except in compliance with the
SpTBXLib License.
A copy of the SpTBXLib License may be found in SpTBXLib-LICENSE.txt or at:
  http://www.silverpointdevelopment.com/sptbxlib/SpTBXLib-LICENSE.htm

Alternatively, the contents of this file may be used under the terms of the
Mozilla Public License Version 1.1 (the "MPL v1.1"), in which case the provisions
of the MPL v1.1 are applicable instead of those in the SpTBXLib License.
A copy of the MPL v1.1 may be found in MPL-LICENSE.txt or at:
  http://www.mozilla.org/MPL/

If you wish to allow use of your version of this file only under the terms of
the MPL v1.1 and not to allow others to use your version of this file under the
SpTBXLib License, indicate your decision by deleting the provisions
above and replace them with the notice and other provisions required by the
MPL v1.1. If you do not delete the provisions above, a recipient may use your
version of this file under either the SpTBXLib License or the MPL v1.1.

Software distributed under the License is distributed on an "AS IS" basis,
WITHOUT WARRANTY OF ANY KIND, either expressed or implied. See the License for
the specific language governing rights and limitations under the License.

The initial developer of this code is Robert Lee.

Requirements:
  - Jordan Russell's Toolbar 2000
    http://www.jrsoftware.org

==============================================================================}

interface

{$BOOLEVAL OFF}   // Unit depends on short-circuit boolean evaluation
{$IF CompilerVersion >= 25} // for Delphi XE4 and up
  {$LEGACYIFEND ON} // requires $IF to be terminated with $ENDIF instead of $IFEND
{$IFEND}

uses
  Windows, Messages, Classes, SysUtils, Controls, Graphics, ImgList, Forms,
  Menus, StdCtrls, ExtCtrls,
  SpTBXSkins, SpTBXItem, SpTBXEditors;

type
  TSpTBXColorListBoxStyle = (
    clbsStandardColors,   // First 16 RGBI colors + 4 additional extended colors
    clbsSystemColors,     // System colors
    clbsIncludeNone,      // Include clNone color
    clbsIncludeDefault,   // Include clDefault color
    clbsCustomColor,      // First color is customizable
    clbsPrettyNames,      // Draw text with pretty names
    clbsNoneAsTransparent // Draw clNone as transparent color
  );
  TSpTBXColorListBoxStyles = set of TSpTBXColorListBoxStyle;

  { TSpTBXColorEditButton }

  TSpTBXColorEditButton = class(TSpTBXEditButton)
  private
    FSelectedColor: TColor;
    procedure SetSelectedColor(const Value: TColor);
  protected
    function DoDrawItem(ACanvas: TCanvas; ARect: TRect; const PaintStage: TSpTBXPaintStage): Boolean; override;
    function GetInternalDropDownMenu: TPopupMenu; override;
  public
    property SelectedColor: TColor read FSelectedColor write SetSelectedColor;
  end;

  { TSpTBXColorEdit }

  TSpTBXColorEdit = class(TSpTBXEdit)
  private
    FColorButton: TSpTBXColorEditButton;
    FSelectedFormat: TSpTBXColorTextType;
    FOnSelectedColorChanged: TNotifyEvent;
    function GetSelectedColor: TColor;
    procedure SetSelectedColor(const Value: TColor);
    procedure SetSelectedFormat(const Value: TSpTBXColorTextType);
    procedure WMKillFocus(var Message: TWMKillFocus); message WM_KILLFOCUS;
  protected
    procedure KeyPress(var Key: Char); override;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure DoSelectedColorChanged; virtual;
    procedure UpdateTextFromValue;
    procedure UpdateValueFromText(RevertWhenInvalid: Boolean = True);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property ColorButton: TSpTBXColorEditButton read FColorButton;
  published
    property Text stored False;
    property SelectedColor: TColor read GetSelectedColor write SetSelectedColor;
    property SelectedFormat: TSpTBXColorTextType read FSelectedFormat write SetSelectedFormat default cttIdentAndHTML;
    property OnSelectedColorChanged: TNotifyEvent read FOnSelectedColorChanged write FOnSelectedColorChanged;
  end;

  { TSpTBXFontComboBox }

  TSpTBXFontPitch = (ffScalable, ffMonospaced);
  TSpTBXFontPitchSet = set of TSpTBXFontPitch;
  TSpTBXFontType = (ffRaster, ffTrueType, ffOpenType);
  TSpTBXFontTypeSet = set of TSpTBXFontType;

  TSpTBXFontComboBoxPreview = class(TCustomControl)
  private
  protected
    procedure CreateParams(var Params: TCreateParams); override;
    procedure Paint; override;
  public
    constructor Create(AOwner: TComponent); override;
    property Caption;
    property Font;
  end;

  TSpTBXFontComboBox = class(TSpTBXComboBox)
  private
    FFontPreview: Boolean;
    FFontNamePreview: Boolean;
    FMaxMRUItems: Integer;
    FMRUCount: Integer;
    FFilterFontPitch: TSpTBXFontPitchSet;
    FFilterFontType: TSpTBXFontTypeSet;
    FPreviewWindow: TSpTBXFontComboBoxPreview;
    FSelectedFont: TFontName;
    FOnFontPreview: TSpTBXEditGetTextEvent;
    FScaledFontGlyphImgList : TImageList;
    procedure UpdateSelectedFont(AddMRU: Boolean);
    procedure SetFontNamePreview(const Value: Boolean);
    procedure SetSelectedFont(const Value: TFontName);
    procedure SetMaxMRUItems(Value: Integer);
    procedure SetFontPreview(const Value: Boolean);
    procedure SetFilterFontPitch(const Value: TSpTBXFontPitchSet);
    procedure SetFilterFontType(const Value: TSpTBXFontTypeSet);
  protected
    procedure Click; override;
    procedure CloseUp; override;
    procedure DoDrawItem(ACanvas: TCanvas; var ARect: TRect; Index: Integer; const State: TOwnerDrawState;
      const PaintStage: TSpTBXPaintStage; var PaintDefault: Boolean); override;
    procedure DropDown; override;
    procedure ChangeScale(M, D: Integer{$IF CompilerVersion >= 31}; isDpiChange: Boolean{$IFEND}); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure MRUAdd(AFontName: TFontName);
    function MRUDelete(AFontName: TFontName): Boolean;
    property MRUCount: Integer read FMRUCount;
  published
    property Items stored False;
    property AutoDropDownWidth default True;
    property AutoItemHeight default False;
    property FontPreview: Boolean read FFontPreview write SetFontPreview default True;
    property FontNamePreview: Boolean read FFontNamePreview write SetFontNamePreview default True;
    property MaxMRUItems: Integer read FMaxMRUItems write SetMaxMRUItems default 5;
    property SelectedFont: TFontName read FSelectedFont write SetSelectedFont;
    property FilterFontPitch: TSpTBXFontPitchSet read FFilterFontPitch write SetFilterFontPitch default [ffScalable, ffMonospaced];
    property FilterFontType: TSpTBXFontTypeSet read FFilterFontType write SetFilterFontType default [ffRaster, ffTrueType, ffOpenType];
    property OnFontPreview: TSpTBXEditGetTextEvent read FOnFontPreview write FOnFontPreview;
  end;

  { TSpTBXColorListBox }

  TSpTBXColorListBox = class(TSpTBXListBox)
  private
    FItems: Boolean;
    FNeedToPopulate: Boolean;
    FSelectedColor: TColor;
    FStyle: TSpTBXColorListBoxStyles;
    procedure ColorCallBack(const AName: string);
    function GetColorName(Index: Integer): string;
    function GetColor(Index: Integer): TColor;
    function GetSelected: TColor;
    procedure SetSelected(const Value: TColor);
    procedure SetStyle(const Value: TSpTBXColorListBoxStyles);
  protected
    procedure CreateWnd; override;
    procedure Loaded; override;
    procedure DblClick; override;
    procedure KeyPress(var Key: Char); override;
    procedure DoDrawItem(ACanvas: TCanvas; var ARect: TRect; Index: Integer; const State: TOwnerDrawState;
      const PaintStage: TSpTBXPaintStage; var PaintDefault: Boolean); override;
    function PickCustomColor: Boolean;
    procedure PopulateList;
  public
    constructor Create(AOwner: TComponent); override;
    procedure AddColor(AColor: TColor; AColorName: string);
    function ColorCount: Integer;
    property Colors[Index: Integer]: TColor read GetColor;
    property ColorNames[Index: Integer]: string read GetColorName;
  published
    property Items: Boolean read FItems; // Hides inherited Items property
    property Selected: TColor read GetSelected write SetSelected default clBlack;
    property Style: TSpTBXColorListBoxStyles read FStyle write SetStyle default [clbsStandardColors, clbsSystemColors, clbsNoneAsTransparent]; // Hides inherited Style property
  end;

{ Helpers }
procedure SpFillFontNames(ADest: TStrings; FilterFontPitch: TSpTBXFontPitchSet; FilterFontType: TSpTBXFontTypeSet);

{ Painting helpers }
procedure SpDrawCheckeredBackground(ACanvas: TCanvas; ARect: TRect);
procedure SpDrawColorDropDownButton(ACanvas: TCanvas; ARect: TRect; Pushed: Boolean; AColor: TColor; PPIScale: TPPIScale; CheckeredBkgndWhenTransparent: Boolean = True);

implementation

uses
  Dialogs, TB2Common,
  {$IF CompilerVersion >= 24} // for Delphi XE3 and up
  System.Types, System.UITypes,
  {$IFEND}
  Themes, SpTBXFormPopupMenu, SpTBXColorPickerForm;

var
  DefaultColorPickerDropDownMenu: TSpTBXColorEditPopupMenu = nil;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Helpers }
type
  TFontEnumRecord = record
    Strings: TStringList;
    FilterFontPitch: TSpTBXFontPitchSet;
    FilterFontType: TSpTBXFontTypeSet;
  end;
  PFontEnumRecord = ^TFontEnumRecord;

function EnumFontsProc(EnumLogFontExDV: PEnumLogFontExDV; EnumTextMetric: PEnumTextMetric;
  FontType: DWORD; LParam: LPARAM): Integer; stdcall;
var
  S: string;
  GlyphIndex: Integer;
  L: TStringList;
  P: TSpTBXFontPitch;
  T: TSpTBXFontType;
  FilterFontPitch: TSpTBXFontPitchSet;
  FilterFontType: TSpTBXFontTypeSet;
const
  NTM_PS_OPENTYPE = $00020000;
  NTM_TT_OPENTYPE = $00040000;
begin
  Result := 1;

  L := PFontEnumRecord(LParam).Strings;
  FilterFontPitch := PFontEnumRecord(LParam).FilterFontPitch;
  FilterFontType := PFontEnumRecord(LParam).FilterFontType;

  GlyphIndex := 0;
  if ((EnumTextMetric.etmNewTextMetricEx.ntmTm.ntmFlags and NTM_TT_OPENTYPE) = NTM_TT_OPENTYPE) or
     ((EnumTextMetric.etmNewTextMetricEx.ntmTm.ntmFlags and NTM_PS_OPENTYPE) = NTM_PS_OPENTYPE) then
  begin
    GlyphIndex := 2;
    P := ffScalable;
    T := ffOpenType;
  end
  else begin
    P := ffScalable;
    if FontType = RASTER_FONTTYPE then
      T := ffRaster
    else
      T := ffTrueType;
  end;

  if (EnumLogFontExDV.elfEnumLogfontEx.elfLogFont.lfPitchAndFamily and FIXED_PITCH) = FIXED_PITCH then
    P := ffMonospaced;

  if not (P in FilterFontPitch) or not (T in FilterFontType) then
    exit;

  S := EnumLogFontExDV.elfEnumLogfontEx.elfLogFont.lfFaceName;
  if (S[1] <> '@') then
    if (L.Count = 0) or not SameText(S, L[L.Count - 1]) then
      L.AddObject(S, Pointer(GlyphIndex));
end;

procedure SpFillFontNames(ADest: TStrings; FilterFontPitch: TSpTBXFontPitchSet;
  FilterFontType: TSpTBXFontTypeSet);
// This will only work on Windows 2000 and above, more info on:
// https://www.delphipraxis.net/616082-post.html#712587
var
  DC: HDC;
  LFont: TLogFont;
  L: TStringList;
  FontEnumRecord: TFontEnumRecord;
begin
  L := TStringList.Create;
  FontEnumRecord.Strings := L;
  FontEnumRecord.FilterFontPitch := FilterFontPitch;
  FontEnumRecord.FilterFontType := FilterFontType;

  DC := GetDC(0);
  try
    FillChar(LFont, SizeOf(LFont), 0);
    LFont.lfCharset := DEFAULT_CHARSET;
    EnumFontFamiliesEx(DC, LFont, @EnumFontsProc, LPARAM(@FontEnumRecord), 0);
    L.Sort;
    ADest.Assign(L);
  finally
    ReleaseDC(0, DC);
    L.Free;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Painting helpers }

procedure SpDrawCheckeredBackground(ACanvas: TCanvas; ARect: TRect);
// Draws a 2x2 white and silver checkered background
var
  R: TRect;
  I, J, HCount, WCount: Integer;
const
  Size = 2;
begin
  WCount := (ARect.Right - ARect.Left) div Size;
  HCount := (ARect.Bottom - ARect.Top) div Size;
  for J := 0 to HCount do
    for I := 0 to WCount do begin
      R := Bounds(ARect.Left + (I * Size), ARect.Top + (J * Size), Size, Size);
      if R.Right > ARect.Right then
        R.Right := ARect.Right;
      if R.Bottom > ARect.Bottom then
        R.Bottom := ARect.Bottom;
      if (I + J) mod 2 = 0 then
        ACanvas.Brush.Color := clWhite
      else
        ACanvas.Brush.Color := clSilver;
      ACanvas.FillRect(R);
    end;
end;

procedure SpDrawColorDropDownButton(ACanvas: TCanvas; ARect: TRect;
  Pushed: Boolean; AColor: TColor;  PPIScale: TPPIScale;
  CheckeredBkgndWhenTransparent: Boolean);
// Draws a button used for color editboxes
var
  R: TRect;
begin
  R := ARect;

  ACanvas.Brush.Color := clBtnFace;
  ACanvas.FillRect(R);
  if not Pushed then
    SpDrawRectangle(ACanvas, R, 0, clBtnHighlight, clBtnShadow);

  InflateRect(R, -PPIScale(2), -PPIScale(2));
  if (AColor = clNone) and CheckeredBkgndWhenTransparent then begin
    // Draw a checkered background when clNone is used
    SpDrawCheckeredBackground(ACanvas, R);
  end
  else begin
    ACanvas.Brush.Color := AColor;
    ACanvas.FillRect(R);
  end;
  SpDrawRectangle(ACanvas, R, 0, clBtnShadow, clBtnHighlight);

  R := ARect;
  R.Left := R.Right - PPIScale(9);
  R.Top := R.Bottom - PPIScale(7);
  ACanvas.Brush.Color := clBtnFace;
  ACanvas.FillRect(R);
  if Pushed then
    SpDrawRectangle(ACanvas, R, 0, clBtnHighlight, clBtnFace)
  else
    SpDrawRectangle(ACanvas, R, 0, clBtnHighlight, clBtnShadow);
  SpDrawArrow(ACanvas, R.Left + (R.Right - R.Left) div 2, R.Top + (R.Bottom - R.Top) div 2 - PPIScale(1), clBlack, True, False, PPIScale(2));

  R := ARect;
  InflateRect(R, -PPIScale(1), -PPIScale(1));
  SpDrawRectangle(ACanvas, R, 0, clBtnFace, clBtnFace);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXColorEditButton }

function TSpTBXColorEditButton.DoDrawItem(ACanvas: TCanvas; ARect: TRect;
  const PaintStage: TSpTBXPaintStage): Boolean;
begin
  if (PaintStage = pstPrePaint) and not BitmapValid then begin
    Result := True;
    if Assigned(OnDraw) then OnDraw(Self, ACanvas, ARect, PaintStage, Result);
    if Result then
      SpDrawColorDropDownButton(ACanvas, ARect, Pushed, FSelectedColor, PPIScale);
  end
  else
    Result := inherited DoDrawItem(ACanvas, ARect, PaintStage);
end;

function TSpTBXColorEditButton.GetInternalDropDownMenu: TPopupMenu;
begin
  if Assigned(DropDownMenu) then
    Result := DropDownMenu
  else
    Result := DefaultColorPickerDropDownMenu;
end;

procedure TSpTBXColorEditButton.SetSelectedColor(const Value: TColor);
begin
  if FSelectedColor <> Value then begin
    FSelectedColor := Value;
    Invalidate;
    if Owner is TSpTBXColorEdit then begin
      TSpTBXColorEdit(Owner).UpdateTextFromValue;
      TSpTBXColorEdit(Owner).DoSelectedColorChanged;
    end;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXColorEdit }

constructor TSpTBXColorEdit.Create(AOwner: TComponent);
begin
  inherited;

  FSelectedFormat := cttIdentAndHTML;

  FColorButton := TSpTBXColorEditButton.Create(Self);
  FColorButton.Parent := Self;
  FColorButton.FreeNotification(Self);
  FColorButton.Align := alRight;
  FColorButton.Width := 19;
  UpdateEditRect;

  Text := 'clBlack';
end;

destructor TSpTBXColorEdit.Destroy;
begin
  FreeAndNil(FColorButton);
  inherited;
end;

procedure TSpTBXColorEdit.DoSelectedColorChanged;
begin
  if Assigned(FOnSelectedColorChanged) then FOnSelectedColorChanged(Self);
end;

procedure TSpTBXColorEdit.KeyPress(var Key: Char);
begin
  inherited;
  if Key = #13 then begin
    Key := #0;
    UpdateValueFromText;
    if FColorButton.IsDroppedDown and Assigned(ActiveFormPopupMenu) then
      ActiveFormPopupMenu.ClosePopup(False);
  end;
end;

procedure TSpTBXColorEdit.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited;
  if (AComponent = FColorButton) and (Operation = opRemove) then
    FColorButton := nil;
end;

function TSpTBXColorEdit.GetSelectedColor: TColor;
begin
  Result := FColorButton.SelectedColor;
end;

procedure TSpTBXColorEdit.SetSelectedColor(const Value: TColor);
begin
  FColorButton.SelectedColor := Value;
end;

procedure TSpTBXColorEdit.SetSelectedFormat(const Value: TSpTBXColorTextType);
begin
  if FSelectedFormat <> Value then begin
    FSelectedFormat := Value;
    UpdateTextFromValue;
  end;
end;

procedure TSpTBXColorEdit.UpdateTextFromValue;
begin
  Text := SpColorToString(SelectedColor, FSelectedFormat);
  SelStart := Length(Text);
end;

procedure TSpTBXColorEdit.UpdateValueFromText(RevertWhenInvalid: Boolean = True);
var
  PrevValue, NewValue, C: TColor;
begin
  PrevValue := SelectedColor;
  NewValue := SelectedColor;

  // Try to parse the text to get the value
  if SpStringToColor(Trim(Text), C) then
    NewValue := C;

  if RevertWhenInvalid or (NewValue <> PrevValue) then begin
    SetSelectedColor(NewValue);
    UpdateTextFromValue;
  end;
end;

procedure TSpTBXColorEdit.WMKillFocus(var Message: TWMKillFocus);
begin
  inherited;
  UpdateValueFromText;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXFontComboBoxPreview }

constructor TSpTBXFontComboBoxPreview.Create(AOwner: TComponent);
begin
  inherited;

  Visible := False;
  SetBounds(0, 0, 0, 0);
  Color := StyleServices.GetSystemColor(clWindow);
  Font.Color := StyleServices.GetSystemColor(clWindowText);
end;

procedure TSpTBXFontComboBoxPreview.CreateParams(var Params: TCreateParams);
const
  CS_DROPSHADOW = $00020000;
begin
  inherited;
  with Params do begin
    Style := (Style and not (WS_CHILD or WS_GROUP or WS_TABSTOP)) or WS_POPUP;
    ExStyle := ExStyle or WS_EX_TOPMOST or WS_EX_TOOLWINDOW;
    WindowClass.Style := WindowClass.Style or CS_SAVEBITS;
    if IsWindowsXP then
      WindowClass.Style := WindowClass.Style or CS_DROPSHADOW;
  end;
end;

procedure TSpTBXFontComboBoxPreview.Paint;
var
  R: TRect;
begin
  inherited;
  R := ClientRect;
  Canvas.Font.Assign(Font);
  SpDrawXPText(Canvas, Caption, R, DT_SINGLELINE or DT_CENTER or DT_VCENTER);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXFontComboBox }

constructor TSpTBXFontComboBox.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FFontNamePreview := True;
  FFontPreview := True;
  FMaxMRUItems := 5;
  FMRUCount := 0;
  AutoItemHeight := False;
  AutoDropDownWidth := True;
  FScaledFontGlyphImgList := TImageList.CreateSize(12, 12);
  FScaledFontGlyphImgList.ResInstLoad(HInstance, rtBitmap, 'SPTBXTRUETYPE', clFuchsia);
  FScaledFontGlyphImgList.ResInstLoad(HInstance, rtBitmap, 'SPTBXOPENTYPE', clFuchsia);
  FFilterFontPitch := [ffScalable, ffMonospaced];
  FFilterFontType := [ffRaster, ffTrueType, ffOpenType];
  ItemHeight := 23;
end;

destructor TSpTBXFontComboBox.Destroy;
begin
  FreeAndNil(FScaledFontGlyphImgList);
  FreeAndNil(FPreviewWindow);
  inherited;
end;

procedure TSpTBXFontComboBox.ChangeScale(M, D: Integer{$IF CompilerVersion >= 31}; isDpiChange: Boolean{$IFEND});
var
  I: Integer;
begin
  inherited;

  // Load scaled glyphs
  I := MulDiv(12, M, D); // Original size of the glyph is 12 (scaled to 18 and 24)
  if I >= 24 then I := 24
  else if I >= 18 then I := 18
  else I := 12;
  if I <> FScaledFontGlyphImgList.Width then begin
    FScaledFontGlyphImgList.Clear;
    FScaledFontGlyphImgList.SetSize(I, I);
    case I of
      12: begin  // 96 PPI, 1x
            FScaledFontGlyphImgList.ResInstLoad(HInstance, rtBitmap, 'SPTBXTRUETYPE', clFuchsia);
            FScaledFontGlyphImgList.ResInstLoad(HInstance, rtBitmap, 'SPTBXOPENTYPE', clFuchsia);
          end;
      18: begin  // 144 PPI, 1.5x
            FScaledFontGlyphImgList.ResInstLoad(HInstance, rtBitmap, 'SPTBXTRUETYPE15x', clFuchsia);
            FScaledFontGlyphImgList.ResInstLoad(HInstance, rtBitmap, 'SPTBXOPENTYPE15x', clFuchsia);
          end;
      24: begin  // 192 PPI, 2x
            FScaledFontGlyphImgList.ResInstLoad(HInstance, rtBitmap, 'SPTBXTRUETYPE20x', clFuchsia);
            FScaledFontGlyphImgList.ResInstLoad(HInstance, rtBitmap, 'SPTBXOPENTYPE20x', clFuchsia);
          end;
    end;
  end;
end;

procedure TSpTBXFontComboBox.Click;
begin
  UpdateSelectedFont(False);
  inherited;
end;

procedure TSpTBXFontComboBox.CloseUp;
begin
  UpdateSelectedFont(True);
  inherited;
  FreeAndNil(FPreviewWindow);
end;

procedure TSpTBXFontComboBox.DropDown;
var
  W: Integer;
  P: TPoint;
  Sz: TSize;
  S: string;
begin
  if Items.Count <= 0 then
    SpFillFontNames(Items, FilterFontPitch, FFilterFontType);

  inherited;

  W := SendMessage(Handle, CB_GETDROPPEDWIDTH, 0, 0) + PPIScale(80); // Add more margin
  SendMessage(Handle, CB_SETDROPPEDWIDTH, W, 0);

  if FFontPreview then begin
    S := 'AaBbYyZz';
    FPreviewWindow := TSpTBXFontComboBoxPreview.Create(Self);
    FPreviewWindow.ParentWindow := Application.Handle;
    FPreviewWindow.Font.Size := 14;

    if Assigned(FOnFontPreview) then FOnFontPreview(Self, S);

    FPreviewWindow.Caption := S;
    Sz := SpGetControlTextSize(FPreviewWindow, FPreviewWindow.Font, S);
    Inc(Sz.cx, PPIScale(100));
    Inc(Sz.cy, PPIScale(20));

    W := SendMessage(Handle, CB_GETDROPPEDWIDTH, 0, 0);
    P := Parent.ClientToScreen(Point(Left, Top));

    if P.X + W + Sz.cx > Screen.Width then
      Dec(P.X, Sz.cx)
    else
      Inc(P.X, W);
    if P.Y + Height + Sz.cy > Screen.Height then
      Dec(P.Y, Sz.cy)
    else
      Inc(P.Y, Height);
    FPreviewWindow.SetBounds(P.X, P.Y, Sz.cx, Sz.cy);
    FPreviewWindow.Visible := True;
    ShowWindow(FPreviewWindow.Handle, SW_SHOWNA);
  end;
end;

procedure TSpTBXFontComboBox.DoDrawItem(ACanvas: TCanvas; var ARect: TRect;
  Index: Integer; const State: TOwnerDrawState;
  const PaintStage: TSpTBXPaintStage; var PaintDefault: Boolean);
var
  Flags, ImageIndex: Integer;
  R: TRect;
const
  Spacing = 4;
begin
  inherited DoDrawItem(ACanvas, ARect, Index, State, PaintStage, PaintDefault);
  if (PaintStage = pstPrePaint) and PaintDefault then begin
    PaintDefault := False;  // Override the default painting
    // Draw the item glyph if the font is TrueType/OpenType
    R := ARect;
    R.Left := Spacing;
    R.Top := R.Top + ((R.Bottom - R.Top) - FScaledFontGlyphImgList.Height) div 2;
    ImageIndex := Integer(Items.Objects[Index]) - 1;
    if ImageIndex > -1 then
      FScaledFontGlyphImgList.Draw(ACanvas, R.Left, R.Top, ImageIndex);

    // Draw the item text
    R := ARect;
    R.Left := Spacing + FScaledFontGlyphImgList.Width + Spacing;
    if FFontNamePreview then
      ACanvas.Font.Name := Items[Index];
    Flags := DrawTextBiDiModeFlags(DT_SINGLELINE or DT_VCENTER or DT_NOPREFIX);
    SpDrawXPText(ACanvas, Items[Index], R, Flags);

    // Draw the MRU separator line
    if FMaxMRUItems > 0 then begin
      if Index = MRUCount - 1 then
        SpDrawLine(ACanvas, ARect.Left, ARect.Bottom - 1, ARect.Right, ARect.Bottom - 1, $C0C0C0);
      if Index = MRUCount then
        SpDrawLine(ACanvas, ARect.Left, ARect.Top, ARect.Right, ARect.Top, $C0C0C0);
    end;

    // Update the Font of the PreviewWindow
    if Assigned(FPreviewWindow) and (odSelected in State) then begin
      FPreviewWindow.Font.Name := Items[Index];
      FPreviewWindow.Invalidate;
    end;
  end;
end;

procedure TSpTBXFontComboBox.MRUAdd(AFontName: TFontName);
var
  I, AFontNameIndex: Integer;
begin
  AFontNameIndex := Items.IndexOf(AFontName);
  if (AFontNameIndex > -1) and (FMRUCount < Items.Count) and (FMaxMRUItems > 0) then begin
    // Exit if it's already on the list
    for I := 0 to FMRUCount - 1 do
      if SameText(AFontName, Items[I]) then Exit;

    // Add the font to the top and delete the last MRU item if necessary
    Items.InsertObject(0, Items[AFontNameIndex], Items.Objects[AFontNameIndex]);
    if FMRUCount >= FMaxMRUItems then
      Items.Delete(FMRUCount)
    else
      Inc(FMRUCount);
  end;
end;

function TSpTBXFontComboBox.MRUDelete(AFontName: TFontName): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := 0 to FMRUCount - 1 do
    if SameText(AFontName, Items[I]) then begin
      Items.Delete(I);
      Result := True;
      Break;
    end;
end;

procedure TSpTBXFontComboBox.SetMaxMRUItems(Value: Integer);
begin
  if Value < 0 then Value := 0;
  if FMaxMRUItems <> Value then begin
    FMaxMRUItems := Value;
    while FMRUCount > FMaxMRUItems do begin
      Items.Delete(FMRUCount);
      Dec(FMRUCount);
    end;
  end;
end;

procedure TSpTBXFontComboBox.SetFilterFontPitch(const Value: TSpTBXFontPitchSet);
begin
  if FFilterFontPitch <> Value then begin
    FFilterFontPitch := Value;
    Items.Clear;
  end;
end;

procedure TSpTBXFontComboBox.SetFilterFontType(const Value: TSpTBXFontTypeSet);
begin
  if FFilterFontType <> Value then begin
    FFilterFontType := Value;
    Items.Clear;
  end;
end;

procedure TSpTBXFontComboBox.SetFontNamePreview(const Value: Boolean);
begin
  if FFontNamePreview <> Value then begin
    FFontNamePreview := Value;
    Invalidate;
  end;
end;

procedure TSpTBXFontComboBox.SetFontPreview(const Value: Boolean);
begin
  FFontPreview := Value;
end;

procedure TSpTBXFontComboBox.SetSelectedFont(const Value: TFontName);
var
  I: Integer;
begin
  I := Items.IndexOf(Value);
  if ItemIndex <> I then
    ItemIndex := I;
  UpdateSelectedFont(True);

  // If the Value is not valid clear the text and call the click events
  if I = -1 then begin
    Click;
    Select;
  end;
end;

procedure TSpTBXFontComboBox.UpdateSelectedFont(AddMRU: Boolean);
var
  I: Integer;
begin
  I := ItemIndex;
  if I > -1 then begin
    FSelectedFont := Items[I];
    if [csDesigning, csLoading] * ComponentState = [] then
      if AddMRU then
        MRUAdd(FSelectedFont);
  end
  else
    FSelectedFont := '';
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXColorListBox }

constructor TSpTBXColorListBox.Create(AOwner: TComponent);
begin
  inherited;
  FStyle := [clbsStandardColors, clbsSystemColors, clbsNoneAsTransparent];
  FSelectedColor := clBlack;
  PopulateList;
end;

procedure TSpTBXColorListBox.CreateWnd;
begin
  inherited CreateWnd;
  if FNeedToPopulate then
    PopulateList;
end;

procedure TSpTBXColorListBox.DblClick;
begin
  inherited;
  if ItemIndex = 0 then
    PickCustomColor;
end;

procedure TSpTBXColorListBox.Loaded;
begin
  inherited;
  Selected := FSelectedColor;
  if FNeedToPopulate then
    PopulateList;
end;

procedure TSpTBXColorListBox.AddColor(AColor: TColor; AColorName: string);
begin
  (inherited Items).AddObject(AColorName, TObject(AColor));
end;

procedure TSpTBXColorListBox.ColorCallBack(const AName: string);
var
  I, LStart: Integer;
  LColor: TColor;
  LName: string;
begin
  LColor := StringToColor(AName);
  if clbsPrettyNames in Style then begin
    if AName = 'clBtnFace' then LName := 'Button Face'
    else if AName = 'clBtnHighlight' then LName := 'Button Highlight'
    else if AName = 'clBtnShadow' then LName := 'Button Shadow'
    else if AName = 'clBtnText' then LName := 'Button Text'
    else if AName = 'clMedGray' then LName := 'Medium Gray'
    else if AName = 'clInfoBk' then LName := 'Info Background'
    else if AName = 'cl3DDkShadow' then LName := '3D Dark Shadow'
    else if AName = 'cl3DLight' then LName := '3D Light'
    else begin
      // Remove the 'cl' prefix and separate the words with a space
      if Copy(AName, 1, 2) = 'cl' then
        LStart := 3
      else
        LStart := 1;
      LName := '';
      for I := LStart to Length(AName) do begin
        case AName[I] of
          'A'..'Z':
            if LName <> '' then
              LName := LName + ' ';
        end;
        LName := LName + AName[I];
      end;
    end;
  end
  else
    LName := AName;

  AddColor(LColor, LName);
end;

function TSpTBXColorListBox.ColorCount: Integer;
begin
  Result := (inherited Items).Count;
end;

procedure TSpTBXColorListBox.DoDrawItem(ACanvas: TCanvas; var ARect: TRect;
  Index: Integer; const State: TOwnerDrawState;
  const PaintStage: TSpTBXPaintStage; var PaintDefault: Boolean);
var
  R: TRect;
  SavedBrushColor: TColor;
begin
  if PaintStage = pstPrePaint then begin
    // Paint the color glyphs
    R := ARect;
    R.Right := R.Left + PPIScale(16 + 5);
    ARect.Left := R.Right + PPIScale(1);
    inherited DoDrawItem(ACanvas, ARect, Index, State, PaintStage, PaintDefault);
    if PaintDefault then begin
      SavedBrushColor := ACanvas.Brush.Color;
      try
        InflateRect(R, -PPIScale(1), -PPIScale(1));

        ACanvas.Brush.Color := Colors[Index];
        if (ACanvas.Brush.Color = clNone) and (clbsNoneAsTransparent in Style) then
          SpDrawCheckeredBackground(ACanvas, R)
        else
          ACanvas.FillRect(R);

        if odSelected in State then
          ACanvas.Brush.Color := clWhite
        else
          ACanvas.Brush.Color := clBlack;
        ACanvas.FrameRect(R);
      finally
        ACanvas.Brush.Color := SavedBrushColor;
      end;
    end;
  end
  else
    inherited DoDrawItem(ACanvas, ARect, Index, State, PaintStage, PaintDefault);
end;

function TSpTBXColorListBox.GetColor(Index: Integer): TColor;
begin
  Result := TColor((inherited Items).Objects[Index]);
end;

function TSpTBXColorListBox.GetColorName(Index: Integer): string;
begin
  Result := (inherited Items)[Index];
end;

function TSpTBXColorListBox.GetSelected: TColor;
begin
  if HandleAllocated then
    if ItemIndex <> -1 then
      Result := Colors[ItemIndex]
    else
      Result := clNone
  else
    Result := FSelectedColor;
end;

procedure TSpTBXColorListBox.KeyPress(var Key: Char);
begin
  inherited KeyPress(Key);
  if (clbsCustomColor in Style) and (Key = #13) and (ItemIndex = 0) then begin
    Key := #0;
    PickCustomColor;
  end;
end;

function TSpTBXColorListBox.PickCustomColor: Boolean;
var
  LColor: TColor;
  CD: TColorDialog;
begin
  Result := False;
  if not (clbsCustomColor in Style) then Exit;

  CD := TColorDialog.Create(nil);
  try
    LColor := ColorToRGB(Colors[0]);
    CD.Color := LColor;
    CD.CustomColors.Text := Format('ColorA=%.8x', [LColor]);
    CD.Options := CD.Options + [cdFullOpen];
    Result := CD.Execute;
    if Result then begin
      (inherited Items).Objects[0] := TObject(CD.Color);
      Invalidate;
    end;
  finally
    CD.Free;
  end;
end;

procedure TSpTBXColorListBox.PopulateList;

  procedure DeleteRange(const AMin, AMax: Integer);
  var
    I: Integer;
  begin
    for I := AMax downto AMin do
      (inherited Items).Delete(I);
  end;

  procedure DeleteColor(const AColor: TColor);
  var
    I: Integer;
  begin
    I := (inherited Items).IndexOfObject(TObject(AColor));
    if I <> -1 then
      (inherited Items).Delete(I);
  end;

var
  LSelectedColor, LCustomColor: TColor;
begin
  if HandleAllocated and not (csLoading in ComponentState) then begin
    (inherited Items).BeginUpdate;
    try
      // Get the custom color
      LCustomColor := clBlack;
      if (clbsCustomColor in Style) and ((inherited Items).Count > 0) then
        LCustomColor := TColor((inherited Items).Objects[0]);

      LSelectedColor := FSelectedColor;
      (inherited Items).Clear;

      GetColorValues(ColorCallBack);
      if not (clbsIncludeNone in Style) then
        DeleteColor(clNone);
      if not (clbsIncludeDefault in Style) then
        DeleteColor(clDefault);
      if not (clbsSystemColors in Style) then
        DeleteRange(StandardColorsCount + ExtendedColorsCount, (inherited Items).Count - 1);
      if not (clbsStandardColors in Style) then
        DeleteRange(0, StandardColorsCount + ExtendedColorsCount - 1);

      if clbsCustomColor in Style then
        (inherited Items).InsertObject(0, '...', TObject(LCustomColor));

      Selected := LSelectedColor;
    finally
      (inherited Items).EndUpdate;
      FNeedToPopulate := False;
    end;
  end
  else
    FNeedToPopulate := True;
end;

procedure TSpTBXColorListBox.SetSelected(const Value: TColor);
var
  I, J: Integer;
begin
  if HandleAllocated then begin
    I := (inherited Items).IndexOfObject(TObject(Value));
    if (I = -1) and (clbsCustomColor in Style) then begin
      // If the color is not on the list add it and select it
      (inherited Items).Objects[0] := TObject(Value);
      I := 0;
    end
    else
      if (clbsCustomColor in Style) and (I = 0) then begin
        // Try to find the color on the list before selecting the 1st custom color
        for J := 1 to (inherited Items).Count - 1 do
          if (inherited Items).Objects[J] = TObject(Value) then begin
            I := J;
            Break;
          end;
      end;

    if (ItemIndex = 0) and (I = 0) then
      Invalidate
    else
      ItemIndex := I;
  end;

  FSelectedColor := Value;
end;

procedure TSpTBXColorListBox.SetStyle(const Value: TSpTBXColorListBoxStyles);
begin
  if Style <> Value then begin
    FStyle := Value;
    Enabled := ([clbsStandardColors, clbsSystemColors, clbsCustomColor] * FStyle) <> [];
    PopulateList;
    if (ColorCount > 0) and (ItemIndex = -1) then
      ItemIndex := 0;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Stock Objects }

procedure InitializeStock;
begin
  Screen.Cursors[crSpTBXEyeDropper] := LoadCursor(HInstance, 'CZEYEDROPPER');
  DefaultColorPickerDropDownMenu := TSpTBXColorEditPopupMenu.Create(nil);
end;

procedure FinalizeStock;
begin
  FreeAndNil(DefaultColorPickerDropDownMenu);
end;

initialization
  InitializeStock;

finalization
  FinalizeStock;

end.

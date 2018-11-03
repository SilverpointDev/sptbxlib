unit SpTBXMDIMRU;

{==============================================================================
Version 2.5.4

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

Development notes:
  - All the theme changes and adjustments are marked with '[Theme-Change]'.

==============================================================================}

interface

{$BOOLEVAL OFF}   // Unit depends on short-circuit boolean evaluation
{$IF CompilerVersion >= 25} // for Delphi XE4 and up
  {$LEGACYIFEND ON} // XE4 and up requires $IF to be terminated with $ENDIF instead of $IFEND
{$IFEND}

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, ImgList, IniFiles, TB2Item, TB2Toolbar, SpTBXSkins, SpTBXItem;

type
  TSpTBXMDIButtonsItem = class;

  TSpTBXMRUListClickEvent = procedure(Sender: TObject; const Filename: string) of object;

  { TSpTBXMDIHandler }

  TSpTBXMDIHandler = class(TComponent)
  private
    FButtonsItem: TSpTBXMDIButtonsItem;
    FSystemMenuItem: TSpTBXSystemMenuItem;
    FToolbar: TTBCustomToolbar;
    procedure SetToolbar(Value: TTBCustomToolbar);
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property Toolbar: TTBCustomToolbar read FToolbar write SetToolbar;
  end;

  { TSpTBXMDIButtonsItem: should only be used by TSpTBXMDIHandler }

  TSpTBXMDIButtonsItem = class(TTBCustomItem)
  private
    FMinimizeItem, FRestoreItem, FCloseItem: TSpTBXItem;
    procedure DrawItem(Sender: TObject; ACanvas: TCanvas;
      ARect: TRect; ItemInfo: TSpTBXMenuItemInfo; const PaintStage: TSpTBXPaintStage;
      var PaintDefault: Boolean);
    procedure DrawItemImage(Sender: TObject; ACanvas: TCanvas;
      State: TSpTBXSkinStatesType; const PaintStage: TSpTBXPaintStage;
      var AImageList: TCustomImageList; var AImageIndex: Integer;
      var ARect: TRect; var PaintDefault: Boolean);
    procedure InvalidateSystemMenuItem;
    procedure ItemClick(Sender: TObject);
    procedure UpdateState(W: HWND; Maximized: Boolean);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

  { TSpTBXMDIWindowItem }

  TSpTBXMDIWindowItem = class(TTBCustomItem)
  private
    FForm: TForm;
    FOnUpdate: TNotifyEvent;
    FWindowMenu: TMenuItem;
    procedure ItemClick(Sender: TObject);
    procedure SetForm(AForm: TForm);
  protected
    procedure EnabledChanged; override;
    procedure GetChildren(Proc: TGetChildProc; Root: TComponent); override;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
  public
    constructor Create(AOwner: TComponent); override;
    procedure InitiateAction; override;
  published
    property Enabled;
    property OnUpdate: TNotifyEvent read FOnUpdate write FOnUpdate;
  end;

  { TSpTBXMRUListItem }

  TSpTBXMRUListItem = class(TTBCustomItem)
  private
    FMaxItems: Integer;
    FOnClick: TSpTBXMRUListClickEvent;
    FHidePathExtension: Boolean;
    procedure ClickHandler(Sender: TObject);
    procedure SetHidePathExtension(const Value: Boolean);
    procedure SetMaxItems(const Value: Integer);
  public
    constructor Create(AOwner: TComponent); override;
    procedure GetMRUFilenames(MRUFilenames: TStrings);
    function IndexOfMRU(Filename: string): Integer;
    function MRUAdd(Filename: string): Integer;
    function MRUClick(Filename: string): Boolean;
    procedure MRURemove(Filename: string);
    procedure MRUUpdateCaptions;
    procedure LoadFromIni(Ini: TCustomIniFile; const Section: string);
    procedure SaveToIni(Ini: TCustomIniFile; const Section: string);
  published
    property HidePathExtension: Boolean read FHidePathExtension write SetHidePathExtension default True;
    property MaxItems: Integer read FMaxItems write SetMaxItems default 4;
    property OnClick: TSpTBXMRUListClickEvent read FOnClick write FOnClick;
  end;

  { TSpTBXMRUItem: should only be used by TSpTBXMRUListItem }

  TSpTBXMRUItem = class(TSpTBXCustomItem)
  private
    FMRUString: string;
  public
    property MRUString: string read FMRUString write FMRUString;
  end;

implementation

uses
  Themes, UxTheme,
  TB2Common, TB2Consts;

type
  TTBCustomToolbarAccess = class(TTBCustomToolbar);
  TTBCustomItemAccess = class(TTBCustomItem);

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXMDIHandler }

constructor TSpTBXMDIHandler.Create(AOwner: TComponent);
begin
  inherited;
  FSystemMenuItem := TSpTBXSystemMenuItem.Create(Self);
  FSystemMenuItem.MDISystemMenu := True;
  FButtonsItem := TSpTBXMDIButtonsItem.Create(Self);
end;

destructor TSpTBXMDIHandler.Destroy;
begin
  SetToolbar(nil);
  inherited;
end;

procedure TSpTBXMDIHandler.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited;
  if (AComponent = FToolbar) and (Operation = opRemove) then
    Toolbar := nil;
end;

procedure TSpTBXMDIHandler.SetToolbar(Value: TTBCustomToolbar);
var
  Rebuild: Boolean;
begin
  if FToolbar <> Value then begin
    if Assigned(FToolbar) then begin
      Rebuild := False;
      if TTBCustomToolbarAccess(FToolbar).FMDIButtonsItem = FButtonsItem then begin
        TTBCustomToolbarAccess(FToolbar).FMDIButtonsItem := nil;
        Rebuild := True;
      end;
      if TTBCustomToolbarAccess(FToolbar).FMDISystemMenuItem = FSystemMenuItem then begin
        TTBCustomToolbarAccess(FToolbar).FMDISystemMenuItem := nil;
        Rebuild := True;
      end;
      if Rebuild and Assigned(FToolbar.View) then
        FToolbar.View.RecreateAllViewers;
    end;
    FToolbar := Value;
    if Assigned(Value) then begin
      Value.FreeNotification(Self);
      TTBCustomToolbarAccess(Value).FMDIButtonsItem := FButtonsItem;
      TTBCustomToolbarAccess(Value).FMDISystemMenuItem := FSystemMenuItem;
      Value.View.RecreateAllViewers;
    end;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXMDIButtonsItem }

var
  CBTHookHandle: HHOOK;
  MDIButtonsItems: TList;

function WindowIsMDIChild(W: HWND): Boolean;
var
  I: Integer;
  MainForm, ChildForm: TForm;
begin
  MainForm := Application.MainForm;
  if Assigned(MainForm) then
    for I := 0 to MainForm.MDIChildCount-1 do begin
      ChildForm := MainForm.MDIChildren[I];
      if ChildForm.HandleAllocated and (ChildForm.Handle = W) then begin
        Result := True;
        Exit;
      end;
    end;
  Result := False;
end;

function CBTHook(Code: Integer; WParam: WPARAM; LParam: LPARAM): LRESULT;
stdcall;
var
  Maximizing: Boolean;
  WindowPlacement: TWindowPlacement;
  I: Integer;
begin
  case Code of
    HCBT_SETFOCUS: begin
        if WindowIsMDIChild(HWND(WParam)) and Assigned(MDIButtonsItems) then begin
          for I := 0 to MDIButtonsItems.Count-1 do
            TSpTBXMDIButtonsItem(MDIButtonsItems[I]).InvalidateSystemMenuItem;
        end;
      end;
    HCBT_MINMAX: begin
        if WindowIsMDIChild(HWND(WParam)) and Assigned(MDIButtonsItems) and
           (LParam in [SW_SHOWNORMAL, SW_SHOWMAXIMIZED, SW_MINIMIZE, SW_RESTORE]) then begin
          Maximizing := (LParam = SW_MAXIMIZE);
          if (LParam = SW_RESTORE) and not IsZoomed(HWND(WParam)) then begin
            WindowPlacement.length := SizeOf(WindowPlacement);
            GetWindowPlacement(HWND(WParam), @WindowPlacement);
            Maximizing := (WindowPlacement.flags and WPF_RESTORETOMAXIMIZED <> 0);
          end;
          for I := 0 to MDIButtonsItems.Count-1 do
            TSpTBXMDIButtonsItem(MDIButtonsItems[I]).UpdateState(HWND(WParam), Maximizing);
        end;
      end;
    HCBT_DESTROYWND: begin
        if WindowIsMDIChild(HWND(WParam)) and Assigned(MDIButtonsItems) then begin
          for I := 0 to MDIButtonsItems.Count-1 do
            TSpTBXMDIButtonsItem(MDIButtonsItems[I]).UpdateState(HWND(WParam), False);
        end;
      end;
  end;
  Result := CallNextHookEx(CBTHookHandle, Code, WParam, LParam);
end;

constructor TSpTBXMDIButtonsItem.Create(AOwner: TComponent);

  function CreateItem(ImageIndex: Integer): TSpTBXItem;
  var
    A: TTBCustomItemAccess;
  begin
    Result := TSpTBXItem.Create(Self);
    A := TTBCustomItemAccess(Result);
    A.ItemStyle := A.ItemStyle + [tbisRightAlign];
    Result.Images := MDIButtonsImgList;
    Result.ImageIndex := ImageIndex;
    Result.CustomWidth := 17;
    Result.OnClick := ItemClick;
    Result.OnDrawItem := DrawItem;
    Result.OnDrawImage := DrawItemImage;
  end;

begin
  inherited;
  ItemStyle := ItemStyle + [tbisEmbeddedGroup];

  FMinimizeItem := CreateItem(2);
  FRestoreItem := CreateItem(3);
  FCloseItem := CreateItem(0);

  Add(FMinimizeItem);
  Add(FRestoreItem);
  Add(FCloseItem);

  UpdateState(0, False);
  AddToList(MDIButtonsItems, Self);
  if CBTHookHandle = 0 then
    CBTHookHandle := SetWindowsHookEx(WH_CBT, CBTHook, 0, GetCurrentThreadId);
end;

destructor TSpTBXMDIButtonsItem.Destroy;
begin
  RemoveFromList(MDIButtonsItems, Self);
  if (MDIButtonsItems = nil) and (CBTHookHandle <> 0) then begin
    UnhookWindowsHookEx(CBTHookHandle);
    CBTHookHandle := 0;
  end;
  inherited;
end;

procedure TSpTBXMDIButtonsItem.DrawItem(Sender: TObject; ACanvas: TCanvas;
  ARect: TRect; ItemInfo: TSpTBXMenuItemInfo;
  const PaintStage: TSpTBXPaintStage; var PaintDefault: Boolean);
begin
  // [Theme-Change]
  // Don't draw the items background if the Default theme is used
  if (PaintStage = pstPrePaint) and SkinManager.IsDefaultSkin then
    PaintDefault := False;
end;

procedure TSpTBXMDIButtonsItem.DrawItemImage(Sender: TObject; ACanvas: TCanvas;
  State: TSpTBXSkinStatesType; const PaintStage: TSpTBXPaintStage;
  var AImageList: TCustomImageList; var AImageIndex: Integer; var ARect: TRect;
  var PaintDefault: Boolean);
const
   ButtonIndexFlags: array[0..3] of Integer = (DFCS_CAPTIONCLOSE, DFCS_CAPTIONMAX, DFCS_CAPTIONMIN, DFCS_CAPTIONRESTORE);
   NoneFlags: array[TSpTBXSkinStatesType] of Integer = (0, DFCS_INACTIVE, 0, DFCS_PUSHED, DFCS_PUSHED, DFCS_PUSHED);
   XPPart: array [0..3] of Integer = (WP_MDICLOSEBUTTON, WP_MAXBUTTON, WP_MDIMINBUTTON, WP_MDIRESTOREBUTTON);
   XPFlags: array[TSpTBXSkinStatesType] of Integer = (CBS_NORMAL, CBS_DISABLED, CBS_HOT, CBS_PUSHED, CBS_PUSHED, CBS_PUSHED);
begin
  if (PaintStage = pstPrePaint) and (AImageList = MDIButtonsImgList) and
    (AImageIndex >= 0) and (AImageIndex <= 3) then
  begin
    case SkinManager.GetSkinType of
      sknNone:
        begin
          PaintDefault := False;
          DrawFrameControl(ACanvas.Handle, ARect, DFC_CAPTION, ButtonIndexFlags[AImageIndex] or NoneFlags[State]);
        end;
      sknWindows:
        begin
          PaintDefault := False;
          DrawThemeBackground(SpTBXThemeServices.Theme[teWindow], ACanvas.Handle, XPPart[AImageIndex], XPFlags[State], ARect, nil);
        end;
    end;
  end;
end;

procedure TSpTBXMDIButtonsItem.UpdateState(W: HWND; Maximized: Boolean);
var
  HasMaxChild, VisibilityChanged: Boolean;

  procedure UpdateVisibleEnabled(const Item: TTBCustomItem; const AEnabled: Boolean);
  begin
    if (Item.Visible <> HasMaxChild) or (Item.Enabled <> AEnabled) then begin
      Item.Visible := HasMaxChild;
      Item.Enabled := AEnabled;
      VisibilityChanged := True;
    end;
  end;

var
  MainForm, ActiveMDIChild, ChildForm: TForm;
  I: Integer;
begin
  HasMaxChild := False;
  MainForm := Application.MainForm;
  ActiveMDIChild := nil;
  if Assigned(MainForm) then begin
    for I := 0 to MainForm.MDIChildCount - 1 do begin
      ChildForm := MainForm.MDIChildren[I];
      if ChildForm.HandleAllocated and
         (((ChildForm.Handle = W) and Maximized) or
          ((ChildForm.Handle <> W) and IsZoomed(ChildForm.Handle))) then begin
        HasMaxChild := True;
        Break;
      end;
    end;
    ActiveMDIChild := MainForm.ActiveMDIChild;
  end;

  VisibilityChanged := False;
  UpdateVisibleEnabled(TSpTBXMDIHandler(Owner).FSystemMenuItem, True);
  UpdateVisibleEnabled(FMinimizeItem, (ActiveMDIChild = nil) or (GetWindowLong(ActiveMDIChild.Handle, GWL_STYLE) and WS_MINIMIZEBOX <> 0));
  UpdateVisibleEnabled(FRestoreItem, True);
  UpdateVisibleEnabled(FCloseItem, True);

  if VisibilityChanged and Assigned((Owner as TSpTBXMDIHandler).FToolbar) then begin
    TSpTBXMDIHandler(Owner).FToolbar.View.InvalidatePositions;
    TSpTBXMDIHandler(Owner).FToolbar.View.TryValidatePositions;
  end;
end;

procedure TSpTBXMDIButtonsItem.ItemClick(Sender: TObject);
var
  MainForm, ChildForm: TForm;
  Cmd: WPARAM;
  SendTo: HWND;
begin
  MainForm := Application.MainForm;
  if Assigned(MainForm) then begin
    ChildForm := MainForm.ActiveMDIChild;
    if Assigned(ChildForm) then begin
      // Make sure we send the message to a maximized window
      SendTo := ChildForm.Handle;
      while not IsZoomed(SendTo) do begin
        SendTo := GetWindow(SendTo, GW_HWNDNEXT);
        if SendTo = 0 then Exit;
      end;

      if Sender = FRestoreItem then
        Cmd := SC_RESTORE
      else if Sender = FCloseItem then
        Cmd := SC_CLOSE
      else
        Cmd := SC_MINIMIZE;
      SendMessage(SendTo, WM_SYSCOMMAND, Cmd, 0); //GetMessagePos);
    end;
  end;
end;

procedure TSpTBXMDIButtonsItem.InvalidateSystemMenuItem;
var
  View: TTBView;
begin
  if Assigned((Owner as TSpTBXMDIHandler).FToolbar) then begin
    View := TSpTBXMDIHandler(Owner).FToolbar.View;
    View.Invalidate(View.Find(TSpTBXMDIHandler(Owner).FSystemMenuItem));
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXMDIWindowItem }

constructor TSpTBXMDIWindowItem.Create(AOwner: TComponent);
var
  Form: TForm;
begin
  inherited;
  ItemStyle := ItemStyle + [tbisEmbeddedGroup];
  Caption := STBMDIWindowItemDefCaption;
  FWindowMenu := TMenuItem.Create(Self);

  if not (csDesigning in ComponentState) then begin
    { Need to set WindowMenu before MDI children are created. Otherwise the
      list incorrectly shows the first 9 child windows, even if window 10+ is
      active. }
    Form := Application.MainForm;
    if (Form = nil) and (Screen.FormCount > 0) then
      Form := Screen.Forms[0];
    SetForm(Form);
  end;
end;

procedure TSpTBXMDIWindowItem.GetChildren(Proc: TGetChildProc; Root: TComponent);
begin
  // Do nothing
end;

procedure TSpTBXMDIWindowItem.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited;
  if (Operation = opRemove) and (AComponent = FForm) then
    SetForm(nil);
end;

procedure TSpTBXMDIWindowItem.SetForm(AForm: TForm);
begin
  if FForm <> AForm then begin
    if Assigned(FForm) and (FForm.WindowMenu = FWindowMenu) then
      FForm.WindowMenu := nil;
    FForm := AForm;
    if Assigned(FForm) then
      FForm.FreeNotification(Self);
  end;
  if Assigned(FForm) then
    FForm.WindowMenu := FWindowMenu;
end;

procedure TSpTBXMDIWindowItem.EnabledChanged;
var
  I: Integer;
begin
  inherited;
  for I := 0 to Count-1 do
    Items[I].Enabled := Enabled;
end;

procedure TSpTBXMDIWindowItem.InitiateAction;
var
  MainForm: TForm;
  I: Integer;
  M: HMENU;
  Item: TSpTBXItem;
  ItemCount: Integer;
  Buf: array[0..1023] of Char;
begin
  inherited;
  if csDesigning in ComponentState then
    Exit;
  MainForm := Application.MainForm;
  if Assigned(MainForm) then
    SetForm(MainForm);
  if FForm = nil then
    Exit;
  if FForm.ClientHandle <> 0 then
    { This is needed, otherwise windows selected on the More Windows dialog
      don't move back into the list }
    SendMessage(FForm.ClientHandle, WM_MDIREFRESHMENU, 0, 0);
  M := FWindowMenu.Handle;
  ItemCount := GetMenuItemCount(M) - 1;
  if ItemCount < 0 then
    ItemCount := 0;
  while Count < ItemCount do begin
    Item := TSpTBXItem.Create(Self);
    Item.Enabled := Enabled;
    Item.OnClick := ItemClick;
    Add(Item);
  end;
  while Count > ItemCount do
    Items[Count - 1].Free;
  for I := 0 to ItemCount - 1 do begin
    Item := TSpTBXItem(Items[I]);
    Item.Tag := GetMenuItemID(M, I+1);
    if GetMenuString(M, I+1, Buf, SizeOf(Buf), MF_BYPOSITION) = 0 then
      Buf[0] := #0;
    Item.Caption := Buf;
    Item.Checked := GetMenuState(M, I+1, MF_BYPOSITION) and MF_CHECKED <> 0;
  end;
  if Assigned(FOnUpdate) then
    FOnUpdate(Self);
end;

procedure TSpTBXMDIWindowItem.ItemClick(Sender: TObject);
var
  Form: TForm;
begin
  Form := Application.MainForm;
  if Assigned(Form) then
    PostMessage(Form.Handle, WM_COMMAND, TTBCustomItem(Sender).Tag, 0);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXMRUListItem }

procedure TSpTBXMRUListItem.ClickHandler(Sender: TObject);
var
  I: Integer;
  A: TSpTBXMRUItem;
begin
  if Sender is TSpTBXMRUItem then begin
    A := TSpTBXMRUItem(Sender);
    I := IndexOf(A);
    if I > 0 then
      Move(I, 0);
    MRUUpdateCaptions;
    if Assigned(FOnClick) then FOnClick(Self, A.MRUString);
  end;
end;

constructor TSpTBXMRUListItem.Create(AOwner: TComponent);
begin
  inherited;
  ItemStyle := ItemStyle + [tbisEmbeddedGroup];
  Caption := STBMRUListItemDefCaption;
  Options := Options + [tboShowHint];
  FMaxItems := 4;
  FHidePathExtension := True;
end;

procedure TSpTBXMRUListItem.GetMRUFilenames(MRUFilenames: TStrings);
var
  I: Integer;
begin
  MRUFilenames.Clear;
  for I := 0 to Count - 1 do
    if Items[I] is TSpTBXMRUItem then
      MRUFilenames.Add(TSpTBXMRUItem(Items[I]).MRUString);
end;

function TSpTBXMRUListItem.IndexOfMRU(Filename: string): Integer;
var
  I: Integer;
begin
  Result := -1;
  for I := 0 to Count - 1 do
    if Items[I] is TSpTBXMRUItem then
      if SameText(TSpTBXMRUItem(Items[I]).MRUString, Filename) then begin
        Result := I;
        Break;
      end;
end;

function TSpTBXMRUListItem.MRUAdd(Filename: string): Integer;
var
  A: TSpTBXMRUItem;
  I: Integer;
begin
  Result := -1;
  I := IndexOfMRU(Filename);
  if I > -1 then begin
    // If Filename is already in the MRU list, move it to the top
    Move(I, 0);
    MRUUpdateCaptions;
  end
  else begin
    // Add the new Filename, if it exceeds the MaxItems limit delete the bottom items
    A := TSpTBXMRUItem.Create(Self);
    A.MRUString := Filename;
    A.Hint := Filename;
    A.OnClick := ClickHandler;
    Insert(0, A);
    while Count > FMaxItems do
      Items[Count - 1].Free;
    MRUUpdateCaptions;
    Result := 0;
  end;
end;

function TSpTBXMRUListItem.MRUClick(Filename: string): Boolean;
var
  I: Integer;
begin
  Result := False;
  I := IndexOfMRU(Filename);
  if I > -1 then begin
    Items[I].Click;
    Result := True;
  end;
end;

procedure TSpTBXMRUListItem.MRURemove(Filename: string);
var
  I: Integer;
begin
  I := IndexOfMRU(Filename);
  if I > -1 then begin
    Items[I].Free;
    MRUUpdateCaptions;
  end;
end;

procedure TSpTBXMRUListItem.MRUUpdateCaptions;
var
  I: Integer;
  A: TSpTBXMRUItem;
  S: string;
begin
  for I := 0 to Count - 1 do
    if Items[I] is TSpTBXMRUItem then begin
      A := TSpTBXMRUItem(Items[I]);
      S := A.MRUString;
      if FHidePathExtension then begin
        S := ExtractFileName(S);
      end;
      A.Caption := '&' + IntToStr(I + 1) + ' ' + S;
    end;
end;

procedure TSpTBXMRUListItem.LoadFromIni(Ini: TCustomIniFile; const Section: string);
var
  I: Integer;
  S: string;
begin
  Clear;
  for I := FMaxItems downto 1 do begin
    S := Ini.ReadString(Section, IntToStr(I), '');
    if S <> '' then
      MRUAdd(S);
  end;
end;

procedure TSpTBXMRUListItem.SaveToIni(Ini: TCustomIniFile; const Section: string);
var
  I: Integer;
  A: TSpTBXMRUItem;
  S: string;
  L: TStringList;
begin
  // If the IniFile doesn't exist create it with Unicode encoding
  // otherwise we can't write Unicode strings.
  if (Ini is TIniFile) and not FileExists(Ini.FileName) then begin
    L := TStringList.Create;
    try
      L.SaveToFile(Ini.FileName, TEncoding.Unicode);
    finally
      L.Free;
    end;
  end;
  for I := 1 to FMaxItems do begin
    if I <= Count then begin
      A := TSpTBXMRUItem(Items[I - 1]);
      S := A.MRUString;
      Ini.WriteString(Section, IntToStr(I), S);
    end
    else
      Ini.DeleteKey(Section, IntToStr(I));
  end;
end;

procedure TSpTBXMRUListItem.SetHidePathExtension(const Value: Boolean);
begin
  if FHidePathExtension <> Value then begin
    FHidePathExtension := Value;
    MRUUpdateCaptions;
  end;
end;

procedure TSpTBXMRUListItem.SetMaxItems(const Value: Integer);
begin
  if FMaxItems <> Value then begin
    FMaxItems := Value;
    if Count > FMaxItems then begin
      while Count > FMaxItems do
        Items[Count - 1].Free;
      MRUUpdateCaptions;
    end;
  end;
end;

end.

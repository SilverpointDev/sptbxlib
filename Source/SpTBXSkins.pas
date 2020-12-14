unit SpTBXSkins;

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
  - All the Windows and Delphi bugs fixes are marked with '[Bugfix]'.
  - All the theme changes and adjustments are marked with '[Theme-Change]'.
  - DT_END_ELLIPSIS and DT_PATH_ELLIPSIS doesn't work with rotated text
    http://support.microsoft.com/kb/249678
  - Vista theme elements are defined on the Themes unit on Delphi XE2 and up.
    VCL Styles are supported on Delphi XE2 and up.
    For older versions of Delphi we should fill the element details manually.
    When XE2 Styles are used menus and toolbar elements (teMenu and teToolbar)
    are painted by TCustomStyleMenuElements.DrawElement.
    State is passed by TCustomStyle.GetElementDetails(Detail: TThemedMenu) as:
    State = Integer(Detail), Part is not used:
    function TCustomStyle.GetElementDetails(Detail: TThemedMenu): TThemedElementDetails;
    begin
      Result.Element := teMenu;
      Result.Part := 0;
      Result.State := Integer(Detail);
    end;
    All this adjustments are marked with '[Old-Themes]'

==============================================================================}

interface

{$BOOLEVAL OFF}   // Unit depends on short-circuit boolean evaluation
{$IF CompilerVersion >= 25} // for Delphi XE4 and up
  {$LEGACYIFEND ON} // XE4 and up requires $IF to be terminated with $ENDIF instead of $IFEND
{$IFEND}

uses
  Windows, Messages, Classes, SysUtils, Graphics, Controls, StdCtrls,
  ImgList, IniFiles, Types,
  {$IF CompilerVersion >= 25} // for Delphi XE4 and up
  System.UITypes,
  {$IFEND}
  Themes, Styles, Generics.Collections;

resourcestring
  SSpTBXColorNone = 'None';
  SSpTBXColorDefault = 'Default';

const
  WM_SPSKINCHANGE = WM_APP + 2007;   // Skin change notification message

type
  { Skins }

  TSpTBXSkinType = (
    sknNone,         // No themes
    sknWindows,      // Use Windows themes
    sknSkin,         // Use Skins
    sknDelphiStyle   // Use Delphi Custom Styles
  );

  TSpTBXLunaScheme = (
    lusBlue,
    lusMetallic,
    lusGreen,
    lusUnknown
  );

  TSpTBXSkinComponentsType = (
    skncDock,
    skncDockablePanel,
    skncDockablePanelTitleBar,
    skncGutter,
    skncMenuBar,
    skncOpenToolbarItem,
    skncPanel,
    skncPopup,
    skncSeparator,
    skncSplitter,
    skncStatusBar,
    skncStatusBarGrip,
    skncTabBackground,
    skncTabToolbar,
    skncToolbar,
    skncToolbarGrip,
    skncWindow,
    skncWindowTitleBar,

    // Multiple States
    skncMenuBarItem,
    skncMenuItem,
    skncToolbarItem,
    skncButton,
    skncCheckBox,
    skncEditButton,
    skncEditFrame,
    skncHeader,
    skncLabel,
    skncListItem,
    skncProgressBar,
    skncRadioButton,
    skncTab,
    skncTrackBar,
    skncTrackBarButton
  );
  TSpTBXSkinStatesType = (sknsNormal, sknsDisabled, sknsHotTrack, sknsPushed, sknsChecked, sknsCheckedAndHotTrack);

  TSpTBXSkinStatesSet = set of TSpTBXSkinStatesType;

  TSpTBXSkinPartsType = (sknpBody, sknpBorders, sknpText);

  TSpTBXSkinComponentsIdentEntry = record
    Name: string;
    States: TSpTBXSkinStatesSet;
  end;

  TPPIScale = function(Value: Integer): Integer of object;

const
  SpTBXSkinMultiStateComponents: set of TSpTBXSkinComponentsType = [skncMenuBarItem..High(TSpTBXSkinComponentsType)];

  CSpTBXSkinAllStates = [Low(TSpTBXSkinStatesType)..High(TSpTBXSkinStatesType)];
  CSpTBXSkinComponents: array [TSpTBXSkinComponentsType] of TSpTBXSkinComponentsIdentEntry = (
    // Single state Components
    (Name: 'Dock';                  States: [sknsNormal]),
    (Name: 'DockablePanel';         States: [sknsNormal]),
    (Name: 'DockablePanelTitleBar'; States: [sknsNormal]),
    (Name: 'Gutter';                States: [sknsNormal]),
    (Name: 'MenuBar';               States: [sknsNormal]),
    (Name: 'OpenToolbarItem';       States: [sknsNormal]),
    (Name: 'Panel';                 States: [sknsNormal]),
    (Name: 'Popup';                 States: [sknsNormal]),
    (Name: 'Separator';             States: [sknsNormal]),
    (Name: 'Splitter';              States: [sknsNormal]),
    (Name: 'StatusBar';             States: [sknsNormal]),
    (Name: 'StatusBarGrip';         States: [sknsNormal]),
    (Name: 'TabBackground';         States: [sknsNormal]),
    (Name: 'TabToolbar';            States: [sknsNormal]),
    (Name: 'Toolbar';               States: [sknsNormal]),
    (Name: 'ToolbarGrip';           States: [sknsNormal]),
    (Name: 'Window';                States: [sknsNormal]),
    (Name: 'WindowTitleBar';        States: [sknsNormal]),
    // Multi state Components
    (Name: 'MenuBarItem';           States: CSpTBXSkinAllStates),
    (Name: 'MenuItem';              States: CSpTBXSkinAllStates),
    (Name: 'ToolbarItem';           States: CSpTBXSkinAllStates),
    (Name: 'Button';                States: CSpTBXSkinAllStates),
    (Name: 'CheckBox';              States: CSpTBXSkinAllStates),
    (Name: 'EditButton';            States: CSpTBXSkinAllStates),
    (Name: 'EditFrame';             States: [sknsNormal, sknsDisabled, sknsHotTrack]),
    (Name: 'Header';                States: [sknsNormal, sknsDisabled, sknsHotTrack, sknsPushed]),
    (Name: 'Label';                 States: [sknsNormal, sknsDisabled]),
    (Name: 'ListItem';              States: CSpTBXSkinAllStates),
    (Name: 'ProgressBar';           States: [sknsNormal, sknsHotTrack]),
    (Name: 'RadioButton';           States: CSpTBXSkinAllStates),
    (Name: 'Tab';                   States: CSpTBXSkinAllStates),
    (Name: 'TrackBar';              States: [sknsNormal, sknsHotTrack]),
    (Name: 'TrackBarButton';        States: [sknsNormal, sknsPushed])
  );

  SSpTBXSkinStatesString: array [TSpTBXSkinStatesType] of string = ('Normal', 'Disabled', 'HotTrack', 'Pushed', 'Checked', 'CheckedAndHotTrack');
  SSpTBXSkinDisplayStatesString: array [TSpTBXSkinStatesType] of string = ('Normal', 'Disabled', 'Hot', 'Pushed', 'Checked', 'Checked && Hot');  

type
  { Text }

  TSpTextRotationAngle = (
    tra0,                      // No rotation
    tra90,                     // 90 degree rotation
    tra270                     // 270 degree rotation
  );

  TSpTBXTextInfo = record
    Text: string;
    TextAngle: TSpTextRotationAngle;
    TextFlags: Cardinal;
    TextSize: TSize;
    IsCaptionShown: Boolean;
    IsTextRotated: Boolean;
  end;

  TSpGlyphLayout = (
    ghlGlyphLeft,                 // Glyph icon on the left of the caption
    ghlGlyphTop                   // Glyph icon on the top of the caption
  );

  TSpGlowDirection = (
    gldNone,                      // No glow
    gldAll,                       // Glow on Left, Top, Right and Bottom of the text
    gldTopLeft,                   // Glow on Top-Left of the text
    gldBottomRight                // Glow on Bottom-Right of the text
  );

  TSpTBXGlyphPattern = (
    gptClose,
    gptMaximize,
    gptMinimize,
    gptRestore,
    gptToolbarClose,
    gptChevron,
    gptVerticalChevron,
    gptMenuCheckmark,
    gptCheckmark,
    gptMenuRadiomark
  );

  { MenuItem }

  TSpTBXComboPart = (cpNone, cpCombo, cpSplitLeft, cpSplitRight);
  TSpTBXMenuItemMarginsInfo = record
    Margins: TRect;               // MenuItem margins
    GutterSize: Integer;          // Size of the gutter
    LeftCaptionMargin: Integer;   // Left margin of the caption
    RightCaptionMargin: Integer;  // Right margin of the caption
    ImageTextSpace: Integer;      // Space between the Icon and the caption
  end;

  TSpTBXMenuItemInfo = record
    CurrentPPI: Integer;
    Enabled: Boolean;
    HotTrack: Boolean;
    Pushed: Boolean;
    Checked: Boolean;
    HasArrow: Boolean;
    ImageShown: Boolean;
    ImageOrCheckShown: Boolean;
    ImageSize: TSize;
    RightImageSize: TSize;
    IsDesigning: Boolean;
    IsOnMenuBar: Boolean;
    IsOnToolbox: Boolean;
    IsOpen: Boolean;
    IsSplit: Boolean;
    IsSunkenCaption: Boolean;
    IsVertical: Boolean;
    MenuMargins: TSpTBXMenuItemMarginsInfo; // Used only on menu items
    ComboPart: TSpTBXComboPart;
    ComboRect: TRect;
    ComboState: TSpTBXSkinStatesType;
    ToolbarStyle: Boolean;
    State: TSpTBXSkinStatesType;
    SkinType: TSpTBXSkinType;
  end;

  { Colors }

  TSpTBXColorTextType = (
    cttDefault,        // Use color idents (clWhite), if not possible use Delphi format ($FFFFFF)
    cttHTML,           // HTML format (#FFFFFF)
    cttIdentAndHTML    // Use color idents (clWhite), if not possible use HTML format
  );

  { TSpTBXSkinOptions }

  TSpTBXSkinOptionEntry = class(TPersistent)
  private
    FSkinType: Integer;
    FColor1, FColor2, FColor3, FColor4: TColor;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create; virtual;
    procedure Fill(ASkinType: Integer; AColor1, AColor2, AColor3, AColor4: TColor);
    procedure ReadFromString(S: string);
    function WriteToString: string;
    function IsEmpty: Boolean;
    function IsEqual(AOptionEntry: TSpTBXSkinOptionEntry): Boolean;
    procedure Lighten(Amount: Integer);
    procedure Reset;
  published
    property SkinType: Integer read FSkinType write FSkinType;
    property Color1: TColor read FColor1 write FColor1;
    property Color2: TColor read FColor2 write FColor2;
    property Color3: TColor read FColor3 write FColor3;
    property Color4: TColor read FColor4 write FColor4;
  end;

  TSpTBXSkinOptionCategory = class(TPersistent)
  private
    FBody: TSpTBXSkinOptionEntry;
    FBorders: TSpTBXSkinOptionEntry;
    FTextColor: TColor;
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create; virtual;
    destructor Destroy; override;
    function IsEmpty: Boolean;
    procedure Reset;
    procedure LoadFromIni(MemIni: TMemIniFile; Section, Ident: string);
    procedure SaveToIni(MemIni: TMemIniFile; Section, Ident: string);
  published
    property Body: TSpTBXSkinOptionEntry read FBody write FBody;
    property Borders: TSpTBXSkinOptionEntry read FBorders write FBorders;
    property TextColor: TColor read FTextColor write FTextColor;
  end;

  TSpTBXSkinOptions = class(TPersistent)
  private
    FColorBtnFace: TColor;
    FFloatingWindowBorderSize: Integer;
    FOptions: array [TSpTBXSkinComponentsType, TSpTBXSkinStatesType] of TSpTBXSkinOptionCategory;
    FOfficeIcons: Boolean;
    FOfficeMenu: Boolean;
    FOfficeStatusBar: Boolean;
    FSkinAuthor: string;
    FSkinName: string;
    function GetOfficeIcons: Boolean;
    function GetOfficeMenu: Boolean;
    function GetOfficePopup: Boolean;
    function GetOfficeStatusBar: Boolean;
    function GetFloatingWindowBorderSize: Integer;
    procedure SetFloatingWindowBorderSize(const Value: Integer);
  protected
    procedure AssignTo(Dest: TPersistent); override;
    procedure BroadcastChanges;
  public
    constructor Create; virtual;
    destructor Destroy; override;
    procedure CopyOptions(AComponent, ToComponent: TSpTBXSkinComponentsType);
    procedure FillOptions; virtual;
    function Options(Component: TSpTBXSkinComponentsType; State: TSpTBXSkinStatesType): TSpTBXSkinOptionCategory; overload;
    function Options(Component: TSpTBXSkinComponentsType): TSpTBXSkinOptionCategory; overload;
    procedure LoadFromFile(Filename: string);
    procedure LoadFromStrings(L: TStrings); virtual;
    procedure SaveToFile(Filename: string);
    procedure SaveToStrings(L: TStrings); virtual;
    procedure SaveToMemIni(MemIni: TMemIniFile); virtual;
    procedure Reset(ForceResetSkinProperties: Boolean = False);

    // Metrics
    procedure GetMenuItemMargins(ACanvas: TCanvas; ImgSize: Integer; out MarginsInfo: TSpTBXMenuItemMarginsInfo; DPI: Integer); virtual;
    function GetState(Enabled, Pushed, HotTrack, Checked: Boolean): TSpTBXSkinStatesType; overload;
    procedure GetState(State: TSpTBXSkinStatesType; out Enabled, Pushed, HotTrack, Checked: Boolean); overload;
    function GetTextColor(Component: TSpTBXSkinComponentsType; State: TSpTBXSkinStatesType): TColor; virtual;
    function GetThemedElementDetails(Component: TSpTBXSkinComponentsType; Enabled, Pushed, HotTrack, Checked, Focused, Defaulted, Grayed: Boolean; out Details: TThemedElementDetails): Boolean; overload;
    function GetThemedElementDetails(Component: TSpTBXSkinComponentsType; State: TSpTBXSkinStatesType; out Details: TThemedElementDetails): Boolean; overload;
    function GetThemedElementSize(ACanvas: TCanvas; Details: TThemedElementDetails; DPI: Integer): TSize;
    procedure GetThemedElementTextColor(Details: TThemedElementDetails; out AColor: TColor);
    function GetThemedSystemColor(AColor: TColor): TColor;

    // Skin Paint
    procedure PaintBackground(ACanvas: TCanvas; ARect: TRect; Component: TSpTBXSkinComponentsType; State: TSpTBXSkinStatesType; Background, Borders: Boolean; Vertical: Boolean = False; ForceRectBorders: TAnchors = []); virtual;
    procedure PaintThemedElementBackground(ACanvas: TCanvas; ARect: TRect; Details: TThemedElementDetails; DPI: Integer); overload;
    procedure PaintThemedElementBackground(ACanvas: TCanvas; ARect: TRect; Component: TSpTBXSkinComponentsType; State: TSpTBXSkinStatesType; DPI: Integer); overload;
    procedure PaintThemedElementBackground(ACanvas: TCanvas; ARect: TRect; Component: TSpTBXSkinComponentsType; Enabled, Pushed, HotTrack, Checked, Focused, Defaulted, Grayed: Boolean; DPI: Integer); overload;

    // Element Paint
    procedure PaintMenuCheckMark(ACanvas: TCanvas; ARect: TRect; Checked, Grayed: Boolean; State: TSpTBXSkinStatesType; DPI: Integer); virtual;
    procedure PaintMenuRadioMark(ACanvas: TCanvas; ARect: TRect; Checked: Boolean; State: TSpTBXSkinStatesType; DPI: Integer); virtual;
    procedure PaintWindowFrame(ACanvas: TCanvas; ARect: TRect; IsActive, DrawBody: Boolean; BorderSize: Integer = 4); virtual;

    // Properties
    property ColorBtnFace: TColor read FColorBtnFace write FColorBtnFace;
    property FloatingWindowBorderSize: Integer read GetFloatingWindowBorderSize write SetFloatingWindowBorderSize; // Unscaled
    property OfficeIcons: Boolean read GetOfficeIcons write FOfficeIcons;
    property OfficeMenu: Boolean read GetOfficeMenu write FOfficeMenu;
    property OfficePopup: Boolean read GetOfficePopup;
    property OfficeStatusBar: Boolean read GetOfficeStatusBar write FOfficeStatusBar;
    property SkinAuthor: string read FSkinAuthor write FSkinAuthor;
    property SkinName: string read FSkinName write FSkinName;
  end;

  TSpTBXSkinOptionsClass = class of TSpTBXSkinOptions;

  { TSpTBXSkinsList }

  TSpTBXSkinsListEntry = class
  public
    SkinClass: TSpTBXSkinOptionsClass;
    SkinStrings: TStringList;
    destructor Destroy; override;
  end;

  TSpTBXSkinsDictionary = TObjectDictionary<string, TSpTBXSkinsListEntry>;
  TSpTBXSkinsDictionaryPair = TPair<string, TSpTBXSkinsListEntry>;

  { TSpTBXSkinManager }

  TSpTBXSkinManager = class
  private
    FCurrentSkin: TSpTBXSkinOptions;
    FNotifies: TList;
    FSkinsList: TSpTBXSkinsDictionary;
    FOnSkinChange: TNotifyEvent;
    procedure Broadcast;
    function GetCurrentSkinName: string;
    procedure ResetToSystemStyle;
  public
    constructor Create; virtual;
    destructor Destroy; override;

    function GetSkinType: TSpTBXSkinType;
    procedure GetSkinsAndDelphiStyles(SkinsAndStyles: TStrings);
    function IsDefaultSkin: Boolean;
    function IsXPThemesEnabled: Boolean;

    procedure AddSkinNotification(AObject: TObject);
    procedure RemoveSkinNotification(AObject: TObject);
    procedure BroadcastSkinNotification;

    procedure LoadFromFile(Filename: string);
    procedure SaveToFile(Filename: string);

    procedure AddSkin(SkinName: string; SkinClass: TSpTBXSkinOptionsClass); overload;
    procedure AddSkin(SkinOptions: TStrings); overload;
    function AddSkinFromFile(Filename: string): string;
    procedure SetToDefaultSkin;
    procedure SetSkin(SkinName: string);

    // [Old-Themes]
    {$IF CompilerVersion >= 23} // for Delphi XE2 and up
    procedure SetDelphiStyle(StyleName: string);
    function IsValidDelphiStyle(StyleName: string): Boolean;
    {$IFEND}

    property CurrentSkin: TSpTBXSkinOptions read FCurrentSkin;
    property CurrentSkinName: string read GetCurrentSkinName;
    property SkinsList: TSpTBXSkinsDictionary read FSkinsList;
    property OnSkinChange: TNotifyEvent read FOnSkinChange write FOnSkinChange;
  end;

  { TSpTBXSkinSwitcher }

  TSpTBXSkinSwitcher = class(TComponent)
  private
    FOnSkinChange: TNotifyEvent;
    function GetSkin: string;
    procedure SetSkin(const Value: string);
    procedure WMSpSkinChange(var Message: TMessage); message WM_SPSKINCHANGE;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property Skin: string read GetSkin write SetSkin;
    property OnSkinChange: TNotifyEvent read FOnSkinChange write FOnSkinChange;
  end;

  { TSpTBXThemeServices }
  {$IF CompilerVersion >= 23} //for Delphi XE2 and up
  TSpTBXThemeServices = TCustomStyleServices;
  {$ELSE}
  TSpTBXThemeServices = TThemeServices;
  {$IFEND}

  { TSpPrintWindow }
  // Use SpPrintWindow instead of PaintTo as many controls will not render
  // properly (no text on editors, no scrollbars, incorrect borders, etc)
  // http://msdn2.microsoft.com/en-us/library/ms535695.aspx
  TSpPrintWindow = function(Hnd: HWND; HdcBlt: HDC; nFlags: UINT): BOOL; stdcall;

{ Delphi Styles}
function SpStyleGetElementObject(Style: TCustomStyleServices; const ControlName, ElementName: string): TObject;
function SpStyleDrawBitmapElement(Element: TObject; State: TSpTBXSkinStatesType; DC: HDC; const R: TRect; ClipRect: PRect; DPI: Integer): Boolean;
function SpStyleStretchDrawBitmapElement(Element: TObject; State: TSpTBXSkinStatesType; DC: HDC; const R: TRect; ClipRect: PRect; DPI: Integer): Boolean;

{ Themes }
function SpTBXThemeServices: TSpTBXThemeServices;
function SkinManager: TSpTBXSkinManager;
function CurrentSkin: TSpTBXSkinOptions;
function SpGetLunaScheme: TSpTBXLunaScheme;
procedure SpFillGlassRect(ACanvas: TCanvas; ARect: TRect);
function SpIsGlassPainting(AControl: TControl): Boolean;
procedure SpDrawParentBackground(Control: TControl; DC: HDC; R: TRect);

{ WideString helpers }
function SpCreateRotatedFont(DC: HDC; Orientation: Integer = 2700): HFONT;
function SpDrawRotatedText(const DC: HDC; AText: string; var ARect: TRect; const AFormat: Cardinal; RotationAngle: TSpTextRotationAngle = tra270): Integer;
function SpCalcXPText(ACanvas: TCanvas; ARect: TRect; Caption: string; CaptionAlignment: TAlignment; Flags: Cardinal; GlyphSize, RightGlyphSize: TSize; Layout: TSpGlyphLayout; PushedCaption: Boolean; out ACaptionRect, AGlyphRect, ARightGlyphRect: TRect; PPIScale: TPPIScale; RotationAngle: TSpTextRotationAngle = tra0): Integer;
function SpDrawXPGlassText(ACanvas: TCanvas; Caption: string; var ARect: TRect; Flags: Cardinal; CaptionGlowSize: Integer): Integer;
function SpDrawXPText(ACanvas: TCanvas; Caption: string; var ARect: TRect; Flags: Cardinal; CaptionGlow: TSpGlowDirection = gldNone; CaptionGlowColor: TColor = clYellow; RotationAngle: TSpTextRotationAngle = tra0): Integer; overload;
// The Following 2 overloads are not used
function SpDrawXPText(ACanvas: TCanvas; ARect: TRect; Caption: string; CaptionGlow: TSpGlowDirection; CaptionGlowColor: TColor; CaptionAlignment: TAlignment; Flags: Cardinal; GlyphSize: TSize; Layout: TSpGlyphLayout; PushedCaption: Boolean; out ACaptionRect, AGlyphRect: TRect; PPIScale: TPPIScale; RotationAngle: TSpTextRotationAngle = tra0): Integer; overload;
function SpDrawXPText(ACanvas: TCanvas; ARect: TRect; Caption: string; CaptionGlow: TSpGlowDirection; CaptionGlowColor: TColor; CaptionAlignment: TAlignment; Flags: Cardinal; IL: TCustomImageList; ImageIndex: Integer; Layout: TSpGlyphLayout; Enabled, PushedCaption, DisabledIconCorrection: Boolean; out ACaptionRect, AGlyphRect: TRect; PPIScale: TPPIScale; RotationAngle: TSpTextRotationAngle = tra0): Integer; overload;
function SpGetTextSize(DC: HDC; S: string; NoPrefix: Boolean): TSize;
function SpGetControlTextHeight(AControl: TControl; AFont: TFont): Integer;
function SpGetControlTextSize(AControl: TControl; AFont: TFont; S: string): TSize;
function SpStripAccelChars(S: string): string;
function SpStripShortcut(S: string): string;
function SpStripTrailingPunctuation(S: string): string;
function SpRectToString(R: TRect): string;
function SpStringToRect(S: string; out R: TRect): Boolean;

{ Color helpers }
function SpColorToHTML(const Color: TColor): string;
function SpColorToString(const Color: TColor; TextType: TSpTBXColorTextType = cttDefault): string;
function SpStringToColor(S: string; out Color: TColor): Boolean;
procedure SpGetRGB(Color: TColor; out R, G, B: Integer);
function SpRGBToColor(R, G, B: Integer): TColor;
function SpLighten(Color: TColor; Amount: Integer): TColor;
function SpBlendColors(TargetColor, BaseColor: TColor; Percent: Integer): TColor;
function SpMixColors(TargetColor, BaseColor: TColor; Amount: Byte): TColor;

{ Painting helpers }
function SpCenterRect(Parent: TRect; ChildWidth, ChildHeight: Integer): TRect; overload;
function SpCenterRect(Parent, Child: TRect): TRect; overload;
function SpCenterRectHoriz(Parent: TRect; ChildWidth: Integer): TRect;
function SpCenterRectVert(Parent: TRect; ChildHeight: Integer): TRect;
procedure SpFillRect(ACanvas: TCanvas; const ARect: TRect; BrushColor: TColor; PenColor: TColor = clNone);
procedure SpDrawLine(ACanvas: TCanvas; X1, Y1, X2, Y2: Integer; Color: TColor);
procedure SpDrawRectangle(ACanvas: TCanvas; ARect: TRect; CornerSize: Integer; ColorTL, ColorBR: TColor; ColorTLInternal: TColor = clNone; ColorBRInternal: TColor = clNone; ForceRectBorders: TAnchors = []); overload;
procedure SpDrawRectangle(ACanvas: TCanvas; ARect: TRect; CornerSize: Integer; ColorL, ColorT, ColorR, ColorB, InternalColorL, InternalColorT, InternalColorR, InternalColorB: TColor; ForceRectBorders: TAnchors = []); overload;
procedure SpAlphaBlend(SrcDC, DstDC: HDC; SrcR, DstR: TRect; Alpha: Byte; SrcHasAlphaChannel: Boolean = False);
procedure SpPaintTo(WinControl: TWinControl; ACanvas: TCanvas; X, Y: Integer);

{ ImageList painting }
procedure SpDrawIconShadow(ACanvas: TCanvas; const ARect: TRect; ImageList: TCustomImageList; ImageIndex: Integer);
procedure SpDrawImageList(ACanvas: TCanvas; const ARect: TRect; ImageList: TCustomImageList; ImageIndex: Integer; Enabled, DisabledIconCorrection: Boolean);

{ Gradients }
procedure SpGradient(ACanvas: TCanvas; const ARect: TRect; StartPos, EndPos, ChunkSize: Integer; C1, C2: TColor; const Vertical: Boolean);
procedure SpGradientFill(ACanvas: TCanvas; const ARect: TRect; const C1, C2: TColor; const Vertical: Boolean);
procedure SpGradientFillMirror(ACanvas: TCanvas; const ARect: TRect; const C1, C2, C3, C4: TColor; const Vertical: Boolean);
procedure SpGradientFillMirrorTop(ACanvas: TCanvas; const ARect: TRect; const C1, C2, C3, C4: TColor; const Vertical: Boolean);
procedure SpGradientFillGlass(ACanvas: TCanvas; const ARect: TRect; const C1, C2, C3, C4: TColor; const Vertical: Boolean);

{ Element painting }
procedure SpDrawArrow(ACanvas: TCanvas; X, Y: Integer; AColor: TColor; Vertical, Reverse: Boolean; Size: Integer);
procedure SpDrawDropMark(ACanvas: TCanvas; DropMark: TRect);
procedure SpDrawFocusRect(ACanvas: TCanvas; const ARect: TRect);
procedure SpDrawGlyphPattern(ACanvas: TCanvas; ARect: TRect; Pattern: TSpTBXGlyphPattern; PatternColor: TColor; DPI: Integer);
procedure SpDrawXPButton(ACanvas: TCanvas; ARect: TRect; Enabled, Pushed, HotTrack, Checked, Focused, Defaulted: Boolean; DPI: Integer);
procedure SpDrawXPCheckBoxGlyph(ACanvas: TCanvas; ARect: TRect; Enabled: Boolean; State: TCheckBoxState; HotTrack, Pushed: Boolean; DPI: Integer);
procedure SpDrawXPRadioButtonGlyph(ACanvas: TCanvas; ARect: TRect; Enabled: Boolean; Checked, HotTrack, Pushed: Boolean; DPI: Integer);
procedure SpDrawXPEditFrame(ACanvas: TCanvas; ARect: TRect; Enabled, HotTrack, ClipContent, AutoAdjust: Boolean; DPI: Integer); overload;
procedure SpDrawXPEditFrame(AWinControl: TWinControl; HotTracking, AutoAdjust, HideFrame: Boolean; DPI: Integer); overload;
procedure SpDrawXPGrip(ACanvas: TCanvas; ARect: TRect; LoC, HiC: TColor; DPI: Integer);
procedure SpDrawXPHeader(ACanvas: TCanvas; ARect: TRect; HotTrack, Pushed: Boolean; DPI: Integer);
procedure SpDrawXPListItemBackground(ACanvas: TCanvas; ARect: TRect; Selected, HotTrack, Focused: Boolean; ForceRectBorders: Boolean = False; Borders: Boolean = True);

{ Skins painting }
procedure SpPaintSkinBackground(ACanvas: TCanvas; ARect: TRect; SkinOption: TSpTBXSkinOptionCategory; Vertical: Boolean);
procedure SpPaintSkinBorders(ACanvas: TCanvas; ARect: TRect; SkinOption: TSpTBXSkinOptionCategory; ForceRectBorders: TAnchors = []);

{ Misc }
function SpIsWinVistaOrUp: Boolean;
function SpIsWin10OrUp: Boolean;
function SpGetDirectories(Path: string; L: TStringList): Boolean;

{ DPI }
function SpPPIScale(Value, DPI: Integer): Integer;
function SpPPIScaleToDPI(PPIScale: TPPIScale): Integer;
procedure SpDPIResizeBitmap(Bitmap: TBitmap; const NewWidth, NewHeight, DPI: Integer);
procedure SpDPIScaleImageList(const ImageList: TCustomImageList; M, D: Integer);

{ Stock Objects }
var
  StockBitmap: TBitmap;
  SpPrintWindow: TSpPrintWindow = nil;

implementation

uses
  UxTheme, Forms, Math, TypInfo,
  SpTBXDefaultSkins, CommCtrl, Rtti;

const
  ROP_DSPDxax = $00E20746;

type
  TControlAccess = class(TControl);

var
  FInternalSkinManager: TSpTBXSkinManager = nil;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Delphi Styles }

function SpStyleGetElementObject(Style: TCustomStyleServices; const ControlName, ElementName: string): TObject;
// From Vcl.Styles.GetElementObject
// Returns a TObject that is a TSeStyleObject:
//   TSeBitmapObject(TSeStyleObject)
//   TSeButtonObject(TSeBitmapObject)
//   TSeTextObject(TSeStyleObject)
// Use Bitmap Style Designer to browse the different controls and elements
// of a style, and the class types of such elements.
var
  CustomStyle: TCustomStyle;
  SeStyle, SeStyleSource, SeStyleObject: TObject;
begin
  Result := nil;
  if Style is TCustomStyle then
    CustomStyle := Style as TCustomStyle
  else
    Exit;

  // Use RTTI to access fields, properties and methods of structures on StyleAPI.inc

  // Get Vcl.Styles.TCustomStyle.FSource which is a StyleAPI.inc.TSeStyle (actual Delphi style file)
  SeStyle := TRttiContext.Create.GetType(CustomStyle.ClassType).GetField('FSource').GetValue(CustomStyle).AsObject;
  // Get StyleAPI.inc.TSeStyle.FStyleSource which is a StyleAPI.inc.TSeStyleSource (structure that contains all the style options and bitmaps)
  SeStyleSource := TRttiContext.Create.GetType(SeStyle.ClassType).GetField('FStyleSource').GetValue(SeStyle).AsObject;

  // Find the control (TSeStyleObject)
  // Call StyleAPI.inc.TSeStyleSource.GetObjectByName method that returns a StyleAPI.inc.TSeStyleObject
  SeStyleObject := TRttiContext.Create.GetType(SeStyleSource.ClassType).GetMethod('GetObjectByName').Invoke(SeStyleSource, [ControlName]).AsObject;

  // Find the element of the control (TSeStyleObject/TSeBitmapObject/TSeButtonObject)
  // Call StyleAPI.inc.TSeStyleObject.FindObjectByName method that returns a StyleAPI.inc.TSeStyleObject
  if SeStyleObject <> nil then begin
    SeStyleObject := TRttiContext.Create.GetType(SeStyleObject.ClassType).GetMethod('FindObjectByName').Invoke(SeStyleObject, [ElementName]).AsObject;
    if SeStyleObject <> nil then
      Result := SeStyleObject;
  end;
end;

function SpStyleDrawBitmapElement(Element: TObject; State: TSpTBXSkinStatesType;
  DC: HDC; const R: TRect; ClipRect: PRect; DPI: Integer): Boolean;
// From Vcl.Styles.DrawBitmapElement
// Element should be a TSeStyleObject descendant
var
  V1, V2: TValue;
  SeState: TValue; // TSeState = (ssNormal, ssDesign, ssMaximized, ssMinimized, ssRollup, ssHot, ssPressed, ssFocused, ssDisabled)
  scTileStyle: TValue; // TscTileStyle = (tsTile, tsStretch, tsCenter, tsVertCenterStretch, tsVertCenterTile, tsHorzCenterStretch, tsHorzCenterTile);
  I: Int64;
begin
  Result := False;
  if Element <> nil then begin
    // Use RTTI to access fields, properties and methods of structures on StyleAPI.inc
    // Set StyleAPI.inc.TSeStyleObject.BoundsRect
    V1 := V1.From(R);
    TRttiContext.Create.GetType(Element.ClassType).GetProperty('BoundsRect').SetValue(Element, V1);

    // Set StyleAPI.inc.TSeStyleObject.State
    SeState := TRttiContext.Create.GetType(Element.ClassType).GetProperty('State').GetValue(Element);
    case State of
      sknsDisabled: I := 7; // ssDisabled
      sknsHotTrack: I := 5; // ssHot
      sknsPushed: I := 6;   // ssPressed
    else
      I := 0; // ssNormal
    end;
    SeState := SeState.FromOrdinal(SeState.TypeInfo, I);
    TRttiContext.Create.GetType(Element.ClassType).GetProperty('State').SetValue(Element, SeState);

    with TGDIHandleRecall.Create(DC, OBJ_FONT) do
    try
      // Call StyleAPI.inc.TSeStyleObject.Draw
      V1 := V1.From(Canvas);
      if ClipRect <> nil then
        V2 := V2.From(ClipRect^)
      else
        V2 := V2.From(Rect(-1, -1, -1, -1));
      TRttiContext.Create.GetType(Element.ClassType).GetMethod('Draw').Invoke(Element, [V1, V2, DPI]);
    finally
      Free;
    end;
    Result := True;
  end;
end;

function SpStyleStretchDrawBitmapElement(Element: TObject; State: TSpTBXSkinStatesType;
  DC: HDC; const R: TRect; ClipRect: PRect; DPI: Integer): Boolean;
// From Vcl.Styles.DrawBitmapElement
var
  V: TValue;
  scTileStyle: TValue; // TscTileStyle = (tsTile, tsStretch, tsCenter, tsVertCenterStretch, tsVertCenterTile, tsHorzCenterStretch, tsHorzCenterTile);
begin
  Result := False;
  // Use RTTI to access fields, properties and methods of structures on StyleAPI.inc
  // Assume SeStyleObject is TSeBitmapObject
  // Set StyleAPI.inc.TSeBitmapObject.TileStyle to tsStretch,
  // and after painting reset to original value
  scTileStyle := TRttiContext.Create.GetType(Element.ClassType).GetProperty('TileStyle').GetValue(Element);
  try
    V := V.FromOrdinal(scTileStyle.TypeInfo, 1); // tsStretch = 1
    TRttiContext.Create.GetType(Element.ClassType).GetProperty('TileStyle').SetValue(Element, V);
    Result := SpStyleDrawBitmapElement(Element, State, DC, R, ClipRect, DPI);
  finally
    // Reset to original value
    TRttiContext.Create.GetType(Element.ClassType).GetProperty('TileStyle').SetValue(Element, scTileStyle);
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Themes }

function SpTBXThemeServices: TSpTBXThemeServices;
begin
  {$IF CompilerVersion >= 23} //for Delphi XE2 and up
  Result := StyleServices;
  {$ELSE}
  Result := ThemeServices;
  {$IFEND}
end;

function SkinManager: TSpTBXSkinManager;
begin
  if not Assigned(FInternalSkinManager) then
    FInternalSkinManager := TSpTBXSkinManager.Create;
  Result := FInternalSkinManager;
end;

function CurrentSkin: TSpTBXSkinOptions;
begin
  Result := SkinManager.CurrentSkin;
end;

function SpGetLunaScheme: TSpTBXLunaScheme;
const
  MaxChars = 1024;
var
  pszThemeFileName, pszColorBuff, pszSizeBuf: PWideChar;
  S: string;
begin
  Result := lusUnknown;

  if SkinManager.IsXPThemesEnabled then begin
    GetMem(pszThemeFileName, 2 * MaxChars);
    GetMem(pszColorBuff,     2 * MaxChars);
    GetMem(pszSizeBuf,       2 * MaxChars);
    try
      if not Failed(GetCurrentThemeName(pszThemeFileName, MaxChars, pszColorBuff, MaxChars, pszSizeBuf, MaxChars)) then
        if UpperCase(ExtractFileName(pszThemeFileName)) = 'LUNA.MSSTYLES' then begin
          S := UpperCase(pszColorBuff);
          if S = 'NORMALCOLOR' then
            Result := lusBlue
          else if S = 'METALLIC' then
            Result := lusMetallic
          else if S = 'HOMESTEAD' then
            Result := lusGreen;
        end;
    finally
      FreeMem(pszSizeBuf);
      FreeMem(pszColorBuff);
      FreeMem(pszThemeFileName);
    end;
  end;
end;

procedure SpFillGlassRect(ACanvas: TCanvas; ARect: TRect);
var
  MemDC: HDC;
  PaintBuffer: HPAINTBUFFER;
begin
  PaintBuffer := BeginBufferedPaint(ACanvas.Handle, ARect, BPBF_TOPDOWNDIB, nil, MemDC);
  try
    FillRect(MemDC, ARect, ACanvas.Brush.Handle);
    BufferedPaintMakeOpaque(PaintBuffer, @ARect);
  finally
    EndBufferedPaint(PaintBuffer, True);
  end;
end;

function SpIsGlassPainting(AControl: TControl): Boolean;
var
  LParent: TWinControl;
begin
  Result := csGlassPaint in AControl.ControlState;
  if Result then begin
    LParent := AControl.Parent;
    while (LParent <> nil) and not LParent.DoubleBuffered do
      LParent := LParent.Parent;
    Result := (LParent = nil) or not LParent.DoubleBuffered or (LParent is TCustomForm);
  end;
end;

procedure SpDrawParentBackground(Control: TControl; DC: HDC; R: TRect);
var
  Parent: TWinControl;
  Brush: HBRUSH;
begin
  Parent := Control.Parent;
  if Parent = nil then begin
    Brush := CreateSolidBrush(ColorToRGB(clBtnFace));
    Windows.FillRect(DC, R, Brush);
    DeleteObject(Brush);
  end
  else
    if Parent.HandleAllocated then begin
      if not Parent.DoubleBuffered and (Control is TWinControl) and SkinManager.IsXPThemesEnabled then
        UxTheme.DrawThemeParentBackground(TWinControl(Control).Handle, DC, @R)
      else
        Controls.PerformEraseBackground(Control, DC);
    end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ WideString helpers}

function EnumFontsProc(const lplf: TLogFont; const lptm: TTextMetric;
  dwType: DWORD; lpData: LPARAM): Integer; stdcall;
begin
  Boolean(Pointer(lpData)^) := True;
  Result := 0;
end;

function SpCreateRotatedFont(DC: HDC; Orientation: Integer = 2700): HFONT;
var
  LogFont: TLogFont;
  TM: TTextMetric;
  VerticalFontName: array[0..LF_FACESIZE-1] of Char;
  VerticalFontExists: Boolean;
begin
  if GetObject(GetCurrentObject(DC, OBJ_FONT), SizeOf(LogFont),
     @LogFont) = 0 then begin
    { just in case... }
    Result := 0;
    Exit;
  end;
  LogFont.lfEscapement := Orientation;
  LogFont.lfOrientation := Orientation;
  LogFont.lfOutPrecision := OUT_TT_ONLY_PRECIS;  { needed for Win9x }

  { Don't let a random TrueType font be substituted when MS Sans Serif or
    Microsoft Sans Serif are used. On Windows 2000 and later, hard-code Tahoma
    because Arial can't display Japanese or Thai Unicode characters (on Windows
    2000 at least). On earlier versions, hard-code Arial since NT 4.0 doesn't
    ship with Tahoma, and 9x doesn't do Unicode. }
  if (StrIComp(LogFont.lfFaceName, 'MS Sans Serif') = 0) or
     (StrIComp(LogFont.lfFaceName, 'Microsoft Sans Serif') = 0) then begin
    if Win32MajorVersion >= 5 then
      StrPCopy(LogFont.lfFaceName, 'Tahoma')
    else
      StrPCopy(LogFont.lfFaceName, 'Arial');
    { Set lfHeight to the actual height of the current font. This is needed
      to work around a Windows 98 issue: on a clean install of the OS,
      SPI_GETNONCLIENTMETRICS returns -5 for lfSmCaptionFont.lfHeight. This is
      wrong; it should return -11 for an 8 pt font. With normal, unrotated text
      this actually displays correctly, since MS Sans Serif doesn't support
      sizes below 8 pt. However, when we change to a TrueType font like Arial,
      this becomes a problem because it'll actually create a font that small. }
    if GetTextMetrics(DC, TM) then begin
      { If the original height was negative, keep it negative }
      if LogFont.lfHeight <= 0 then
        LogFont.lfHeight := -(TM.tmHeight - TM.tmInternalLeading)
      else
        LogFont.lfHeight := TM.tmHeight;
    end;
  end;

  { Use a vertical font if available so that Asian characters aren't drawn
    sideways }
  if StrLen(LogFont.lfFaceName) < SizeOf(VerticalFontName)-1 then begin
    VerticalFontName[0] := '@';
    StrCopy(@VerticalFontName[1], LogFont.lfFaceName);
    VerticalFontExists := False;
    EnumFonts(DC, VerticalFontName, @EnumFontsProc, @VerticalFontExists);
    if VerticalFontExists then
      StrCopy(LogFont.lfFaceName, VerticalFontName);
  end;

  Result := CreateFontIndirect(LogFont);
end;

function SpDrawRotatedText(const DC: HDC; AText: string; var ARect: TRect; const AFormat: Cardinal; RotationAngle: TSpTextRotationAngle = tra270): Integer;
{ The format flag this function respects are
  DT_CALCRECT, DT_NOPREFIX, DT_HIDEPREFIX, DT_CENTER, DT_END_ELLIPSIS, DT_NOCLIP }
var
  RotatedFont, SaveFont: HFONT;
  TextMetrics: TTextMetric;
  X, Y, P, I, SU, FU, W: Integer;
  SaveAlign: UINT;
  Clip: Boolean;
  Pen, SavePen: HPEN;
  Sz: TSize;
  Orientation: Integer;
begin
  Result := 0;
  if Length(AText) = 0 then Exit;

  Orientation := 0;
  case RotationAngle of
    tra90: Orientation := 900;   // 90 degrees
    tra270: Orientation := 2700; // 270 degrees
  end;
  RotatedFont := SpCreateRotatedFont(DC, Orientation);
  SaveFont := SelectObject(DC, RotatedFont);

  GetTextMetrics(DC, TextMetrics);
  X := ARect.Left + (ARect.Right - ARect.Left - TextMetrics.tmHeight) div 2;

  Clip := AFormat and DT_NOCLIP = 0;

  { Find the index of the character that should be underlined. Delete '&'
    characters from the string. Like DrawText, only the last prefixed character
    will be underlined. }
  P := 0;
  I := 1;
  if AFormat and DT_NOPREFIX = 0 then
    while I <= Length(AText) do
    begin
      if AText[I] = '&' then
      begin
        Delete(AText, I, 1);
        if PWideChar(AText)[I - 1] <> '&' then P := I;
      end;
      Inc(I);
    end;

  if AFormat and DT_END_ELLIPSIS <> 0 then
  begin
    if (Length(AText) > 1) and (SpGetTextSize(DC, AText, False).cx > ARect.Bottom - ARect.Top) then
    begin
      W := ARect.Bottom - ARect.Top;
      if W > 2 then
      begin
        Delete(AText, Length(AText), 1);
        while (Length(AText) > 1) and (SpGetTextSize(DC, AText + '...', False).cx > W) do
          Delete(AText, Length(AText), 1);
      end
      else AText := AText[1];
      if P > Length(AText) then P := 0;
      AText := AText + '...';
    end;
  end;

  Sz := SpGetTextSize(DC, AText, False);
  Result := Sz.cy;

  if AFormat and DT_CALCRECT <> 0 then begin
    ARect.Right := ARect.Left + Sz.cy;
    ARect.Bottom := ARect.Top + Sz.cx;
  end
  else begin
    if AFormat and DT_CENTER <> 0 then
      Y := ARect.Top + (ARect.Bottom - ARect.Top - Sz.cx) div 2
    else
      Y := ARect.Top;

    if Clip then
    begin
      SaveDC(DC);
      with ARect do IntersectClipRect(DC, Left, Top, Right, Bottom);
    end;

    case RotationAngle of
      tra90: SaveAlign := SetTextAlign(DC, TA_RIGHT);
      tra270: SaveAlign := SetTextAlign(DC, TA_BOTTOM);
    else
      SaveAlign := SetTextAlign(DC, TA_LEFT);
    end;

    Windows.TextOut(DC, X, Y, PWideChar(AText), Length(AText));
    SetTextAlign(DC, SaveAlign);

    { Underline }
    if (P > 0) and (AFormat and DT_HIDEPREFIX = 0) then
    begin
      SU := SpGetTextSize(DC, Copy(AText, 1, P - 1), False).cx;
      FU := SU + SpGetTextSize(DC, PWideChar(AText)[P - 1], False).cx;
      Inc(X, TextMetrics.tmDescent - 2);
      Pen := CreatePen(PS_SOLID, 1, GetTextColor(DC));
      SavePen := SelectObject(DC, Pen);
      MoveToEx(DC, X, Y + SU, nil);
      LineTo(DC, X, Y + FU);
      SelectObject(DC, SavePen);
      DeleteObject(Pen);
    end;

    if Clip then RestoreDC(DC, -1);
  end;

  SelectObject(DC, SaveFont);
  DeleteObject(RotatedFont);
end;

function SpCalcXPText(ACanvas: TCanvas; ARect: TRect; Caption: string;
  CaptionAlignment: TAlignment; Flags: Cardinal; GlyphSize, RightGlyphSize: TSize;
  Layout: TSpGlyphLayout; PushedCaption: Boolean; out ACaptionRect, AGlyphRect, ARightGlyphRect: TRect;
  PPIScale: TPPIScale; RotationAngle: TSpTextRotationAngle = tra0): Integer;
var
  R: TRect;
  TextOffset, Spacing, RightSpacing: TPoint;
  CaptionSz: TSize;
begin
  Result := 0;
  ACaptionRect := Rect(0, 0, 0, 0);
  AGlyphRect := Rect(0, 0, 0, 0);
  ARightGlyphRect := Rect(0, 0, 0, 0);
  TextOffset := Point(0, 0);
  Spacing := Point(0, 0);
  RightSpacing := Point(0, 0);
  if (Caption <> '') and (GlyphSize.cx > 0) and (GlyphSize.cy > 0) then
    Spacing := Point(PPIScale(4), PPIScale(1));
  if (Caption <> '') and (RightGlyphSize.cx > 0) and (RightGlyphSize.cy > 0) then
    RightSpacing := Point(PPIScale(4), PPIScale(1));

  Flags := Flags and not DT_CENTER;
  Flags := Flags and not DT_VCENTER;
  if CaptionAlignment = taRightJustify then
    Flags := Flags or DT_RIGHT;

  // DT_END_ELLIPSIS and DT_PATH_ELLIPSIS doesn't work with rotated text
  // http://support.microsoft.com/kb/249678
  // Revert the ARect if the text is rotated, from now on work on horizontal text !!!
  if RotationAngle <> tra0 then
    ARect := Rect(ARect.Top, ARect.Left, ARect.Bottom, ARect.Right);

  // Get the caption size
  if ((Flags and DT_WORDBREAK) <> 0) or ((Flags and DT_END_ELLIPSIS) <> 0) or ((Flags and DT_PATH_ELLIPSIS) <> 0) then begin
    if Layout = ghlGlyphLeft then  // Glyph on left or right side
      R := Rect(0, 0, ARect.Right - ARect.Left - GlyphSize.cx - Spacing.X - RightGlyphSize.cx - RightSpacing.X + PPIScale(2), PPIScale(1))
    else  // Glyph on top
      R := Rect(0, 0, ARect.Right - ARect.Left + PPIScale(2), PPIScale(1));
  end
  else
    R := Rect(0, 0, PPIScale(1), PPIScale(1));

  if (fsBold in ACanvas.Font.Style) and (RotationAngle = tra0) and (((Flags and DT_END_ELLIPSIS) <> 0) or ((Flags and DT_PATH_ELLIPSIS) <> 0)) then begin
    // [Bugfix] Windows bug:
    // When the Font is Bold and DT_END_ELLIPSIS or DT_PATH_ELLIPSIS is used
    // DrawTextW returns an incorrect size if the string is unicode.
    // The R.Right is reduced by 3 which cuts down the string and
    // adds the ellipsis.
    // We have to obtain the real size and check if it fits in the Rect.
    CaptionSz := SpGetTextSize(ACanvas.Handle, Caption, True);
    if CaptionSz.cx <= R.Right then begin
      R := Rect(0, 0, CaptionSz.cx, CaptionSz.cy);
      Result := CaptionSz.cy;
    end;
  end;

  if Result <= 0 then begin
    Result := SpDrawXPText(ACanvas, Caption, R, Flags or DT_CALCRECT, gldNone, clYellow);
    CaptionSz.cx := R.Right;
    CaptionSz.cy := R.Bottom;
  end;

  // ACaptionRect
  if Result > 0 then begin
    R.Top := ARect.Top + (ARect.Bottom - ARect.Top - CaptionSz.cy) div 2; // Vertically centered
    R.Bottom := R.Top + CaptionSz.cy;
    case CaptionAlignment of
      taCenter:
        R.Left := ARect.Left + (ARect.Right - ARect.Left - CaptionSz.cx) div 2; // Horizontally centered
      taLeftJustify:
        R.Left := ARect.Left;
      taRightJustify:
        R.Left := ARect.Right - CaptionSz.cx;
    end;
    R.Right := R.Left + CaptionSz.cx;

    // Since DT_END_ELLIPSIS and DT_PATH_ELLIPSIS doesn't work with rotated text
    // try to fix it by padding the text 8 pixels to the right
    if (RotationAngle <> tra0) and (R.Right + 8 < ARect.Right) then
      if ((Flags and DT_END_ELLIPSIS) <> 0) or ((Flags and DT_PATH_ELLIPSIS) <> 0) then
        R.Right := R.Right + PPIScale(8);

    if PushedCaption then
      OffsetRect(R, PPIScale(1), PPIScale(1));

    ACaptionRect := R;
  end;

  // AGlyphRect
  if (GlyphSize.cx > 0) and (GlyphSize.cy > 0) then begin
    R := ARect;

    // If ghlGlyphTop is used the glyph should be centered
    if Layout = ghlGlyphTop then
      CaptionAlignment := taCenter;

    case CaptionAlignment of
      taCenter:
        begin
          // Total width = Icon + Space + Text
          if Layout = ghlGlyphLeft then begin
            AGlyphRect.Left := R.Left + (R.Right - R.Left - (GlyphSize.cx + Spacing.X + CaptionSz.cx)) div 2;
            TextOffset.X := (GlyphSize.cx + Spacing.X) div 2;
          end
          else
            AGlyphRect.Left := R.Left + (R.Right - R.Left - GlyphSize.cx) div 2;
        end;
      taLeftJustify:
        begin
          AGlyphRect.Left := R.Left;
          TextOffset.X := GlyphSize.cx + Spacing.X;
        end;
      taRightJustify:
        begin
          AGlyphRect.Left := R.Right - GlyphSize.cx;
          TextOffset.X := - Spacing.X - GlyphSize.cx;
        end;
    end;

    if Layout = ghlGlyphLeft then
      AGlyphRect.Top := R.Top + (R.Bottom - R.Top - GlyphSize.cy) div 2
    else begin
      AGlyphRect.Top := R.Top + (R.Bottom - R.Top - (GlyphSize.cy + Spacing.Y + CaptionSz.cy)) div 2;
      Inc(TextOffset.Y, (GlyphSize.cy + Spacing.Y) div 2);
    end;

    AGlyphRect.Right := AGlyphRect.Left + GlyphSize.cx;
    AGlyphRect.Bottom := AGlyphRect.Top + GlyphSize.cy;

    if PushedCaption then
      OffsetRect(AGlyphRect, PPIScale(1), PPIScale(1));
  end;

  // Move the text according to the icon position
  if Result > 0 then
    OffsetRect(ACaptionRect, TextOffset.X, TextOffset.Y);

  // ARightGlyphRect, it's valid only when using taLeftJustify
  if (RightGlyphSize.cx > 0) and (RightGlyphSize.cy > 0) then
    if CaptionAlignment = taLeftJustify then begin
      R := ARect;
      ARightGlyphRect.Left := R.Right - RightGlyphSize.cx;
      ARightGlyphRect.Right := ARightGlyphRect.Left + RightGlyphSize.cx;
      ARightGlyphRect.Top := R.Top + (R.Bottom - R.Top - RightGlyphSize.cy) div 2;
      ARightGlyphRect.Bottom := ARightGlyphRect.Top + RightGlyphSize.cy;
      if (Result > 0) and (ACaptionRect.Right > ARightGlyphRect.Left - RightSpacing.X) then
        ACaptionRect.Right := ARightGlyphRect.Left - RightSpacing.X;
    end;

  // Revert back, normalize when the text is rotated
  if RotationAngle <> tra0 then begin
    ACaptionRect := Rect(ACaptionRect.Top, ACaptionRect.Left, ACaptionRect.Bottom, ACaptionRect.Right);
    AGlyphRect := Rect(AGlyphRect.Top, AGlyphRect.Left, AGlyphRect.Bottom, AGlyphRect.Right);
    ARightGlyphRect := Rect(ARightGlyphRect.Top, ARightGlyphRect.Left, ARightGlyphRect.Bottom, ARightGlyphRect.Right);
  end;
end;

function SpDrawXPGlassText(ACanvas: TCanvas; Caption: string; var ARect: TRect;
  Flags: Cardinal; CaptionGlowSize: Integer): Integer;

  function InternalDraw(C: TCanvas; var R: TRect): Integer;
  var
    Options: TDTTOpts;
  begin
    FillChar(Options, SizeOf(Options), 0);
    Options.dwSize := SizeOf(Options);
    Options.dwFlags := DTT_TEXTCOLOR or DTT_COMPOSITED or DTT_GLOWSIZE;
    if Flags and DT_CALCRECT = DT_CALCRECT then
      Options.dwFlags := Options.dwFlags or DTT_CALCRECT;
    Options.crText := ColorToRGB(C.Font.Color);
    Options.iGlowSize := CaptionGlowSize;
    DrawThemeTextEx(SpTBXThemeServices.Theme[teWindow], C.Handle, WP_CAPTION, CS_ACTIVE,
      PWideChar(WideString(Caption)), Length(Caption), Flags, @R, Options);

    Result := R.Bottom - R.Top;
  end;

var
  B: TBitmap;
  RBitmap: TRect;
begin
  if Flags and DT_CALCRECT = 0 then begin
    B := TBitmap.Create;
    try
      RBitmap := Rect(0, 0, ARect.Right - ARect.Left, ARect.Bottom - ARect.Top);
      B.PixelFormat := pf32bit;
      // Negative Height because DrawThemeTextEx uses top-down DIBs when DTT_COMPOSITED is used
      // If biHeight is positive, the bitmap is a bottom-up DIB and its origin is the lower left corner.
      // If biHeight is negative, the bitmap is a top-down DIB and its origin is the upper left corner.
      B.SetSize(RBitmap.Right, -RBitmap.Bottom);
      B.Canvas.Font.Assign(ACanvas.Font);
      B.Canvas.Brush.Color := clBlack;  // Fill with black to make it transparent on Glass
      B.Canvas.FillRect(RBitmap);

      Result := InternalDraw(B.Canvas, RBitmap);
      SpAlphaBlend(B.Canvas.Handle, ACanvas.Handle, RBitmap, ARect, 255, True);
    finally
      B.Free;
    end;
  end
  else
    Result := InternalDraw(ACanvas, ARect);
end;

function SpDrawXPText(ACanvas: TCanvas; Caption: string; var ARect: TRect;
  Flags: Cardinal; CaptionGlow: TSpGlowDirection = gldNone;
  CaptionGlowColor: TColor = clYellow; RotationAngle: TSpTextRotationAngle = tra0): Integer; overload;

  function IsCalcRect: Boolean;
  begin
    Result := Flags and DT_CALCRECT = DT_CALCRECT;
  end;

  function InternalDraw(var R: TRect): Integer;
  begin
    Result := 0;
    case RotationAngle of
      tra0:
        Result := Windows.DrawText(ACanvas.Handle, PWideChar(Caption), -1, R, Flags);
      tra90, tra270:
        Result := SpDrawRotatedText(ACanvas.Handle, Caption, R, Flags, RotationAngle);
    end;
  end;

var
  BS: TBrushStyle;
  GlowR: TRect;
  C, FC: TColor;
begin
  BS := ACanvas.Brush.Style;
  C := ACanvas.Brush.Color;
  try
    ACanvas.Brush.Style := bsClear;

    if (CaptionGlow <> gldNone) and not IsCalcRect then begin
      FC := ACanvas.Font.Color;
      ACanvas.Font.Color := CaptionGlowColor;
      case CaptionGlow of
        gldAll:
          begin
            GlowR := ARect; OffsetRect(GlowR, 0, -1);
            InternalDraw(GlowR);
            GlowR := ARect; OffsetRect(GlowR, 0, 1);
            InternalDraw(GlowR);
            GlowR := ARect; OffsetRect(GlowR, -1, 0);
            InternalDraw(GlowR);
            GlowR := ARect; OffsetRect(GlowR, 1, 0);
          end;
        gldTopLeft:
          begin
            GlowR := ARect; OffsetRect(GlowR, -1, -1);
            InternalDraw(GlowR);
          end;
        gldBottomRight:
          begin
            GlowR := ARect; OffsetRect(GlowR, 1, 1);
            InternalDraw(GlowR);
          end;
      end;
      ACanvas.Font.Color := FC;
    end;

    Result := InternalDraw(ARect);

    if IsRectEmpty(ARect) then
      Result := 0
    else
      if IsCalcRect then begin
        // [Bugfix] Windows bug:
        // When DT_CALCRECT is used and the font is italic the
        // resulting rect is incorrect
        if fsItalic in ACanvas.Font.Style then
          ARect.Right := ARect.Right + 1 + (ACanvas.Font.Size div 8) * 2;
      end;

  finally
    ACanvas.Brush.Style := BS;
    ACanvas.Brush.Color := C;
  end;
end;

function SpDrawXPText(ACanvas: TCanvas; ARect: TRect; Caption: string;
  CaptionGlow: TSpGlowDirection; CaptionGlowColor: TColor; CaptionAlignment: TAlignment;
  Flags: Cardinal; GlyphSize: TSize; Layout: TSpGlyphLayout; PushedCaption: Boolean;
  out ACaptionRect, AGlyphRect: TRect; PPIScale: TPPIScale;
  RotationAngle: TSpTextRotationAngle = tra0): Integer; overload;
var
  DummyRightGlyphSize: TSize;
  DummyRightGlyphRect: TRect;
begin
  DummyRightGlyphSize.cx := 0;
  DummyRightGlyphSize.cy := 0;
  DummyRightGlyphRect := Rect(0, 0, 0, 0);
  Result := SpCalcXPText(ACanvas, ARect, Caption, CaptionAlignment, Flags, GlyphSize, DummyRightGlyphSize,
    Layout, PushedCaption, ACaptionRect, AGlyphRect, DummyRightGlyphRect, PPIScale, RotationAngle);
  SpDrawXPText(ACanvas, Caption, ACaptionRect, Flags and not DT_CALCRECT, CaptionGlow, CaptionGlowColor, RotationAngle);
end;

function SpDrawXPText(ACanvas: TCanvas; ARect: TRect; Caption: string;
  CaptionGlow: TSpGlowDirection; CaptionGlowColor: TColor; CaptionAlignment: TAlignment;
  Flags: Cardinal; IL: TCustomImageList; ImageIndex: Integer; Layout: TSpGlyphLayout;
  Enabled, PushedCaption, DisabledIconCorrection: Boolean; out ACaptionRect, AGlyphRect: TRect;
  PPIScale: TPPIScale; RotationAngle: TSpTextRotationAngle = tra0): Integer; overload;
var
  GlyphSize, DummyRightGlyphSize: TSize;
  DummyRightGlyphRect: TRect;
begin
  GlyphSize.cx := 0;
  GlyphSize.cy := 0;
  DummyRightGlyphSize.cx := 0;
  DummyRightGlyphSize.cy := 0;
  DummyRightGlyphRect := Rect(0, 0, 0, 0);

  if Assigned(IL) and (ImageIndex > -1) and (ImageIndex < IL.Count) then begin
    GlyphSize.cx := IL.Width;
    GlyphSize.cy := IL.Height;
  end;

  Result := SpCalcXPText(ACanvas, ARect, Caption, CaptionAlignment, Flags, GlyphSize, DummyRightGlyphSize,
    Layout, PushedCaption, ACaptionRect, AGlyphRect, DummyRightGlyphRect, PPIScale, RotationAngle);

  SpDrawXPText(ACanvas, Caption, ACaptionRect, Flags and not DT_CALCRECT, CaptionGlow, CaptionGlowColor, RotationAngle);

  if Assigned(IL) and (ImageIndex > -1) and (ImageIndex < IL.Count) then
    SpDrawImageList(ACanvas, AGlyphRect, IL, ImageIndex, Enabled, DisabledIconCorrection);
end;

function SpGetTextSize(DC: HDC; S: string; NoPrefix: Boolean): TSize;
// Returns the size of the string, if NoPrefix is True, it first removes "&"
// characters as necessary.
// This procedure is 10x faster than using DrawText with the DT_CALCRECT flag
begin
  Result.cx := 0;
  Result.cy := 0;
  if NoPrefix then
    S := SpStripAccelChars(S);
  Windows.GetTextExtentPoint32W(DC, S, Length(S), Result);
end;

function SpGetControlTextHeight(AControl: TControl; AFont: TFont): Integer;
// Returns the control text height based on the font
var
  Sz: TSize;
begin
  Sz := SpGetControlTextSize(AControl, AFont, 'WQqJ');
  Result := Sz.cy;
end;

function SpGetControlTextSize(AControl: TControl; AFont: TFont; S: string): TSize;
// Returns the control text size based on the font
var
  ACanvas: TControlCanvas;
begin
  ACanvas := TControlCanvas.Create;
  try
    ACanvas.Control := AControl;
    ACanvas.Font.Assign(AFont);
    // newpy MulDiv(ACanvas.Font.Height, AControl.CurrentPPI, AFont.PixelsPerInch);
    Result := SpGetTextSize(ACanvas.Handle, S, False);
  finally
    ACanvas.Free;
  end;
end;

function SpStripAccelChars(S: string): string;
var
  I: Integer;
begin
  Result := S;
  I := 1;
  while I <= Length(Result) do begin
    if (Result[I] = '&') and not ((I + 1 <= Length(Result)) and (Result[I + 1] = '&')) then
      System.Delete(Result, I, 1);  // Don't delete double &&
    Inc(I);
  end;
end;

function SpStripShortcut(S: string): string;
var
  P: Integer;
begin
  Result := S;
  P := Pos(#9, Result);
  if P <> 0 then
    SetLength(Result, P - 1);
end;

function SpStripTrailingPunctuation(S: string): string;
// Removes any colon (':') or ellipsis ('...') from the end of S and returns
// the resulting string
var
  L: Integer;
begin
  Result := S;
  L := Length(Result);
  if (L > 1) and (Result[L] = ':') then
    SetLength(Result, L-1)
  else if (L > 3) and (Result[L-2] = '.') and (Result[L-1] = '.') and
     (Result[L] = '.') then
    SetLength(Result, L-3);
end;

function SpRectToString(R: TRect): string;
begin
  Result := Format('%d, %d, %d, %d', [R.Left, R.Top, R.Right, R.Bottom]);
end;

function SpStringToRect(S: string; out R: TRect): Boolean;
var
  L: TStringList;
begin
  Result := False;
  R := Rect(0, 0, 0, 0);
  L := TStringList.Create;
  try
    L.CommaText := S;
    if L.Count = 4 then begin
      R.Left := StrToIntDef(L[0], 0);
      R.Top := StrToIntDef(L[1], 0);
      R.Right := StrToIntDef(L[2], 0);
      R.Bottom := StrToIntDef(L[3], 0);
      Result := True;
    end;
  finally
    L.Free;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Color Helpers }

function SpColorToHTML(const Color: TColor): string;
var
  R: TColorRef;
begin
  R := ColorToRGB(Color);
  Result := Format('#%.2x%.2x%.2x', [GetRValue(R), GetGValue(R), GetBValue(R)]);
end;

function SpColorToString(const Color: TColor; TextType: TSpTBXColorTextType = cttDefault): string;
begin
  case TextType of
    cttDefault:
      Result := ColorToString(Color);
    cttHTML:
      // Use resourcestring only when clNone or clDefault
      if Color = clNone then Result := SSpTBXColorNone
      else
        if Color = clDefault then Result := SSpTBXColorDefault
        else
          Result := SpColorToHTML(Color);
    cttIdentAndHTML:
      begin
        Result := ColorToString(Color);
        if (Length(Result) > 0) and (Result[1] = '$') then
          Result := SpColorToHTML(Color);
      end;
  end;
end;

function SpStringToColor(S: string; out Color: TColor): Boolean;
var
  L, C: Integer;
begin
  Result := False;
  Color := clDefault;
  L := Length(S);
  if L < 2 then Exit;

  // Try to convert clNone and clDefault resourcestring
  if S = SSpTBXColorNone then begin
    Color := clNone;
    Result := True;
    Exit;
  end
  else
    if S = SSpTBXColorDefault then begin
      Color := clDefault;
      Result := True;
      Exit;
    end;

  if (S[1] = '#') and (L = 7) then begin  // HTML format: #FFFFFF
    S[1] := '$';
    if TryStrToInt(S, C) then begin
      S := Format('$00%s%s%s', [Copy(S, 6, 2), Copy(S, 4, 2), Copy(S, 2, 2)]);
      Color := StringToColor(S);
      Result := True;
    end;
  end
  else
    if (S[1] = '$') and (L >= 7) and (L <= 9) then begin  // TColor format: $FFFFFF
      if TryStrToInt(S, C) then begin
        Color := C;
        Result := True;
      end;
    end
    else
      Result := IdentToColor(S, Longint(Color));  // Ident format: clWhite
end;

procedure SpGetRGB(Color: TColor; out R, G, B: Integer);
begin
  Color := ColorToRGB(Color);
  R := GetRValue(Color);
  G := GetGValue(Color);
  B := GetBValue(Color);
end;

function SpRGBToColor(R, G, B: Integer): TColor;
begin
  if R < 0 then R := 0 else if R > 255 then R := 255;
  if G < 0 then G := 0 else if G > 255 then G := 255;
  if B < 0 then B := 0 else if B > 255 then B := 255;
  Result := TColor(RGB(R, G, B));
end;

function SpLighten(Color: TColor; Amount: Integer): TColor;
var
  R, G, B: Integer;
begin
  Color := ColorToRGB(Color);
  R := GetRValue(Color) + Amount;
  G := GetGValue(Color) + Amount;
  B := GetBValue(Color) + Amount;
  Result := SpRGBToColor(R, G, B);
end;

function SpBlendColors(TargetColor, BaseColor: TColor; Percent: Integer): TColor;
// Blend 2 colors with a predefined percent (0..100 or 0..1000)
// If Percent is 0 the result will be BaseColor,
// If Percent is 100 the result will be TargetColor.
// Any other value will return a color between base and target.
// For example if you want to add 70% of yellow ($0000FFFF) to a color:
// NewColor := SpBlendColor($0000FFFF, BaseColor, 70);
// The result will have 70% of yellow and 30% of BaseColor
var
  Percent2, D, F: Integer;
  R, G, B, R2, G2, B2: Integer;
begin
  SpGetRGB(TargetColor, R, G, B);
  SpGetRGB(BaseColor, R2, G2, B2);

  if Percent >= 100 then D := 1000
  else D := 100;
  Percent2 := D - Percent;
  F := D div 2;

  R := (R * Percent + R2 * Percent2 + F) div D;
  G := (G * Percent + G2 * Percent2 + F) div D;
  B := (B * Percent + B2 * Percent2 + F) div D;

  Result := SpRGBToColor(R, G, B);
end;

function SpMixColors(TargetColor, BaseColor: TColor; Amount: Byte): TColor;
// Mix 2 colors with a predefined amount (0..255).
// If Amount is 0 the result will be BaseColor,
// If Amount is 255 the result will be TargetColor.
// Any other value will return a color between base and target.
// For example if you want to add 50% of yellow ($0000FFFF) to a color:
// NewColor := SpMixColors($0000FFFF, BaseColor, 128);
// The result will be BaseColor + 50% of yellow
var
  R1, G1, B1: Integer;
  R2, G2, B2: Integer;
begin
  SpGetRGB(BaseColor, R1, G1, B1);
  SpGetRGB(TargetColor, R2, G2, B2);

  R1 := (R2 - R1) * Amount div 255 + R1;
  G1 := (G2 - G1) * Amount div 255 + G1;
  B1 := (B2 - B1) * Amount div 255 + B1;

  Result := SpRGBToColor(R1, G1, B1);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Painting Helpers }

function SpCenterRect(Parent: TRect; ChildWidth, ChildHeight: Integer): TRect;
begin
  Result.Left := (Parent.Left + Parent.Right - ChildWidth) div 2;
  Result.Top := (Parent.Top + Parent.Bottom - ChildHeight) div 2;
  Result.Right := Result.Left + ChildWidth;
  Result.Bottom := Result.Top + ChildHeight;
end;

function SpCenterRect(Parent, Child: TRect): TRect;
begin
  Result := SpCenterRect(Parent, Child.Right - Child.Left, Child.Bottom - Child.Top);
end;

function SpCenterRectHoriz(Parent: TRect; ChildWidth: Integer): TRect;
begin
  Result.Left := Parent.Left + (Parent.Right - Parent.Left - ChildWidth) div 2;
  Result.Right := Result.Left + ChildWidth;
  Result.Top := Parent.Top;
  Result.Bottom := Parent.Bottom;
end;

function SpCenterRectVert(Parent: TRect; ChildHeight: Integer): TRect;
begin
  Result.Left := Parent.Left;
  Result.Right := Parent.Right;
  Result.Top := Parent.Top + (Parent.Bottom - Parent.Top - ChildHeight) div 2;
  Result.Bottom := Result.Top + ChildHeight;
end;

procedure SpFillRect(ACanvas: TCanvas; const ARect: TRect; BrushColor: TColor; PenColor: TColor = clNone);
var
  C, C2: TColor;
begin
  if BrushColor <> clNone then begin
    C := ACanvas.Brush.Color;
    C2 := ACanvas.Pen.Color;
    ACanvas.Brush.Color := BrushColor;
    ACanvas.Pen.Color := PenColor;
    if PenColor = clNone then
      ACanvas.FillRect(ARect)
    else
      ACanvas.Rectangle(ARect);
    ACanvas.Brush.Color := C;
    ACanvas.Pen.Color := C2;
  end;
end;

procedure SpDrawLine(ACanvas: TCanvas; X1, Y1, X2, Y2: Integer; Color: TColor);
var
  C: TColor;
begin
  if Color <> clNone then begin
    C := ACanvas.Pen.Color;
    ACanvas.Pen.Color := Color;
    ACanvas.MoveTo(X1, Y1);
    ACanvas.LineTo(X2, Y2);
    ACanvas.Pen.Color := C;
  end;
end;

procedure SpDrawRectangle(ACanvas: TCanvas; ARect: TRect; CornerSize: Integer;
  ColorL, ColorT, ColorR, ColorB,
  InternalColorL, InternalColorT, InternalColorR, InternalColorB: TColor;
  ForceRectBorders: TAnchors = []);
// Draws 2 beveled borders.
// CornerSize can be 0, 1 or 2.
// Color: left, top, right, bottom external border color
// InternalColor: left, top, right, bottom internal border color
// ForceRectBorders: forces the borders to be rect
var
  Color: TColor;
  CornerSizeTL, CornerSizeTR, CornerSizeBL, CornerSizeBR: Integer;
begin
  // Do not use SpDPIScale
  ACanvas.Pen.Width := 1;
  Color := ACanvas.Pen.Color;

  if CornerSize < 0 then CornerSize := 0;
  if CornerSize > 2 then CornerSize := 2;
  CornerSizeTL := CornerSize;
  CornerSizeTR := CornerSize;
  CornerSizeBL := CornerSize;
  CornerSizeBR := CornerSize;
  if akLeft in ForceRectBorders then begin
    CornerSizeTL := 0;
    CornerSizeBL := 0;
  end;
  if akRight in ForceRectBorders then begin
    CornerSizeTR := 0;
    CornerSizeBR := 0;
  end;
  if akTop in ForceRectBorders then begin
    CornerSizeTL := 0;
    CornerSizeTR := 0;
  end;
  if akBottom in ForceRectBorders then begin
    CornerSizeBL := 0;
    CornerSizeBR := 0;
  end;

  with ARect do begin
    Dec(Right);
    Dec(Bottom);

    // Internal borders
    InflateRect(ARect, -1, -1);
    if InternalColorL <> clNone then begin
      ACanvas.Pen.Color := InternalColorL;
      ACanvas.PolyLine([Point(Left, Bottom), Point(Left, Top)]);
    end;
    if InternalColorT <> clNone then begin
      ACanvas.Pen.Color := InternalColorT;
      ACanvas.PolyLine([Point(Left, Top), Point(Right, Top)]);
    end;
    if InternalColorR <> clNone then begin
      ACanvas.Pen.Color := InternalColorR;
      ACanvas.PolyLine([Point(Right, Bottom), Point(Right, Top - 1)]);
    end;
    if InternalColorB <> clNone then begin
      ACanvas.Pen.Color := InternalColorB;
      ACanvas.PolyLine([Point(Left, Bottom), Point(Right, Bottom)]);
    end;

    // External borders
    InflateRect(ARect, 1, 1);
    if ColorL <> clNone then begin
      ACanvas.Pen.Color := ColorL;
      ACanvas.PolyLine([
        Point(Left, Bottom - CornerSizeBL),
        Point(Left, Top + CornerSizeTL)
      ]);
    end;
    if ColorT <> clNone then begin
      ACanvas.Pen.Color := ColorT;
      ACanvas.PolyLine([
        Point(Left, Top + CornerSizeTL),
        Point(Left + CornerSizeTL, Top),
        Point(Right - CornerSizeTR, Top),
        Point(Right, Top + CornerSizeTR)
      ]);
    end;
    if ColorR <> clNone then begin
      ACanvas.Pen.Color := ColorR;
      ACanvas.PolyLine([
        Point(Right , Bottom - CornerSizeBR),
        Point(Right, Top - 1 + CornerSizeTR)
      ]);
    end;
    if ColorB <> clNone then begin
      ACanvas.Pen.Color := ColorB;
      ACanvas.PolyLine([
        Point(Left, Bottom - CornerSizeBL),
        Point(Left + CornerSizeBL, Bottom),
        Point(Right - CornerSizeBR, Bottom),
        Point(Right, Bottom - CornerSizeBR)
      ]);
    end;
  end;

  ACanvas.Pen.Color := Color;
end;

procedure SpDrawRectangle(ACanvas: TCanvas; ARect: TRect;
  CornerSize: Integer; ColorTL, ColorBR, ColorTLInternal, ColorBRInternal: TColor;
  ForceRectBorders: TAnchors);
// Draws 2 beveled borders.
// CornerSize can be 0, 1 or 2.
// TLColor, ColorBR: external border color
// InternalTL, ColorBRInternal: internal border color
// ForceRectBorders: forces the borders to be rect
begin
  SpDrawRectangle(ACanvas, ARect, CornerSize,
    ColorTL, ColorTL, ColorBR, ColorBR,
    ColorTLInternal, ColorTLInternal, ColorBRInternal, ColorBRInternal,
    ForceRectBorders);
end;

procedure SpAlphaBlend(SrcDC, DstDC: HDC; SrcR, DstR: TRect; Alpha: Byte;
  SrcHasAlphaChannel: Boolean = False);
// NOTE: AlphaBlend does not work on Windows 95 and Windows NT
var
  BF: TBlendFunction;
begin
  BF.BlendOp := AC_SRC_OVER;
  BF.BlendFlags := 0;
  BF.SourceConstantAlpha := Alpha;
  if SrcHasAlphaChannel then
    BF.AlphaFormat := AC_SRC_ALPHA
  else
    BF.AlphaFormat := 0;
  Windows.AlphaBlend(DstDC, DstR.Left, DstR.Top, DstR.Right - DstR.Left, DstR.Bottom - DstR.Top,
    SrcDC, SrcR.Left, SrcR.Top, SrcR.Right - SrcR.Left, SrcR.Bottom - SrcR.Top, BF);
end;

procedure SpPaintTo(WinControl: TWinControl; ACanvas: TCanvas; X, Y: Integer);
// NOTE: PrintWindow does not work if the control is not visible
var
  B: TBitmap;
  PrevTop: Integer;
begin
  // Use SpPrintWindow instead of PaintTo as many controls will not render
  // properly (no text on editors, no scrollbars, incorrect borders, etc)
  // http://msdn2.microsoft.com/en-us/library/ms535695.aspx
  if Assigned(SpPrintWindow) then begin
    ACanvas.Lock;
    try
      // It doesn't work if the control is not visible !!!
      // Show it and move it offscreen
      if not WinControl.Visible then begin
        PrevTop := WinControl.Top;
        WinControl.Top := 10000;  // Move it offscreen
        WinControl.Visible := True;
        SpPrintWindow(WinControl.Handle, ACanvas.Handle, 0);
        WinControl.Visible := False;
        WinControl.Top := PrevTop;
      end
      else
        SpPrintWindow(WinControl.Handle, ACanvas.Handle, 0);
    finally
      ACanvas.UnLock;
    end;
  end
  else begin
    // If SpPrintWindow is not available use PaintTo
    // If the Control is a Form use GetFormImage instead
    if WinControl is TCustomForm then begin
      B := TCustomForm(WinControl).GetFormImage;
      try
        ACanvas.Draw(X, Y, B);
      finally
        B.Free;
      end;
    end
    else
      WinControl.PaintTo(ACanvas, X, Y);
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ ImageList painting }

procedure SpDrawIconShadow(ACanvas: TCanvas; const ARect: TRect;
  ImageList: TCustomImageList; ImageIndex: Integer);
var
  ImageWidth, ImageHeight: Integer;
  I, J: Integer;
  Src, Dst: ^Cardinal;
  S, C, CBRB, CBG: Cardinal;
  B1, B2: TBitmap;
begin
  ImageWidth := ARect.Right - ARect.Left;
  ImageHeight := ARect.Bottom - ARect.Top;
  with ImageList do
  begin
    if Width < ImageWidth then ImageWidth := Width;
    if Height < ImageHeight then ImageHeight :=  Height;
  end;

  B1 := TBitmap.Create;
  B2 := TBitmap.Create;
  try
    B1.PixelFormat := pf32bit;
    B2.PixelFormat := pf32bit;
    B1.SetSize(ImageWidth, ImageHeight);
    B2.SetSize(ImageWidth, ImageHeight);

    BitBlt(B1.Canvas.Handle, 0, 0, ImageWidth, ImageHeight, ACanvas.Handle, ARect.Left, ARect.Top, SRCCOPY);
    BitBlt(B2.Canvas.Handle, 0, 0, ImageWidth, ImageHeight, ACanvas.Handle, ARect.Left, ARect.Top, SRCCOPY);
    ImageList.Draw(B2.Canvas, 0, 0, ImageIndex, True);

    for J := 0 to ImageHeight - 1 do
    begin
      Src := B2.ScanLine[J];
      Dst := B1.ScanLine[J];
      for I := 0 to ImageWidth - 1 do
      begin
        S := Src^;
        if S <> Dst^ then
        begin
          CBRB := Dst^ and $00FF00FF;
          CBG  := Dst^ and $0000FF00;
          C := ((S and $00FF0000) shr 16 * 29 + (S and $0000FF00) shr 8 * 150 +
            (S and $000000FF) * 76) shr 8;
          C := (C div 3) + (255 - 255 div 3);
          Dst^ := ((CBRB * C and $FF00FF00) or (CBG * C and $00FF0000)) shr 8;
        end;
        Inc(Src);
        Inc(Dst);
      end;
    end;
    BitBlt(ACanvas.Handle, ARect.Left, ARect.Top, ImageWidth, ImageHeight, B1.Canvas.Handle, 0, 0, SRCCOPY);
  finally
    B1.Free;
    B2.Free;
  end;
end;

procedure SpDrawImageList(ACanvas: TCanvas; const ARect: TRect; ImageList: TCustomImageList;
  ImageIndex: Integer; Enabled, DisabledIconCorrection: Boolean);
begin
  if Assigned(ImageList) and (ImageIndex > -1) and (ImageIndex < ImageList.Count) then
    if not Enabled and DisabledIconCorrection then
      SpDrawIconShadow(ACanvas, ARect, ImageList, ImageIndex)
    else
      ImageList.Draw(ACanvas, ARect.Left, ARect.Top, ImageIndex, Enabled);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Gradients }

procedure SpGradient(ACanvas: TCanvas; const ARect: TRect;
  StartPos, EndPos, ChunkSize: Integer; C1, C2: TColor; const Vertical: Boolean);
// StartPos: start position relative to ARect, usually 0
// EndPos: end position relative to ARect, usually ARect.Bottom - ARect.Top
// ChunkSize: size of the chunk of the gradient we need to paint
  procedure SetVertex(var AVertex: TTriVertex; const APoint: TPoint; ARGBColor: DWORD);
  begin
    AVertex.X := APoint.X;
    AVertex.Y := APoint.Y;
    AVertex.Red := MakeWord(0, GetRValue(ARGBColor));
    AVertex.Green := MakeWord(0, GetGValue(ARGBColor));
    AVertex.Blue := MakeWord(0, GetBValue(ARGBColor));
    AVertex.Alpha := MakeWord(0, 255);
  end;
const
  AModesMap: array[Boolean] of DWORD = (GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V);
var
  AVertices: array[0..1] of TTriVertex;
  AGradientRect: TGradientRect;
  ARGBColor1, ARGBColor2: DWORD;
  Vertex: TPoint;
begin
  ARGBColor1 := ColorToRGB(C1);
  ARGBColor2 := ColorToRGB(C2);
  Vertex := ARect.TopLeft;
  if Vertical then
    Vertex.Y := Vertex.Y + StartPos
  else
    Vertex.X := Vertex.X + StartPos;
  SetVertex(AVertices[0], Vertex, ARGBColor1);
  Vertex := ARect.BottomRight;
  if Vertical then
    Vertex.Y := ARect.Top + EndPos + 1
  else
    Vertex.X := ARect.Left + EndPos + 1;
  SetVertex(AVertices[1], Vertex, ARGBColor2);
  AGradientRect.UpperLeft := 0;
  AGradientRect.LowerRight := 1;
  GradientFill(ACanvas.Handle, @AVertices, 2, @AGradientRect, 1, AModesMap[Vertical]);
end;

procedure SpGradientFill(ACanvas: TCanvas; const ARect: TRect;
  const C1, C2: TColor; const Vertical: Boolean);
var
  GSize: Integer;
begin
  if Vertical then
    GSize := (ARect.Bottom - ARect.Top) - 1
  else
    GSize := (ARect.Right - ARect.Left) - 1;

  SpGradient(ACanvas, ARect, 0, GSize, GSize, C1, C2, Vertical);
end;

procedure SpGradientFillMirror(ACanvas: TCanvas; const ARect: TRect;
  const C1, C2, C3, C4: TColor; const Vertical: Boolean);
var
  GSize, ChunkSize, d1, d2: Integer;
begin
  if Vertical then
    GSize := (ARect.Bottom - ARect.Top) - 1
  else
    GSize := (ARect.Right - ARect.Left) - 1;

  ChunkSize := GSize div 2;
  if ChunkSize = 0 then ChunkSize := 1;
  d1 := ChunkSize;
  d2 := GSize;

  SpGradient(ACanvas, ARect, 0, d1, ChunkSize, C1, C2, Vertical);
  SpGradient(ACanvas, ARect, d1, d2, ChunkSize, C3, C4, Vertical);
end;

procedure SpGradientFillMirrorTop(ACanvas: TCanvas; const ARect: TRect;
  const C1, C2, C3, C4: TColor; const Vertical: Boolean);
var
  GSize, d1, d2: Integer;
begin
  if Vertical then
    GSize := (ARect.Bottom - ARect.Top) - 1
  else
    GSize := (ARect.Right - ARect.Left) - 1;

  d1 := GSize div 3;
  d2 := GSize;

  SpGradient(ACanvas, ARect, 0, d1, d1, C1, C2, Vertical);
  SpGradient(ACanvas, ARect, d1, d2, d2 - d1, C3, C4, Vertical);
end;

procedure SpGradientFillGlass(ACanvas: TCanvas; const ARect: TRect;
  const C1, C2, C3, C4: TColor; const Vertical: Boolean);
var
  GSize, ChunkSize, d1, d2, d3: Integer;
begin
  if Vertical then
    GSize := (ARect.Bottom - ARect.Top) - 1
  else
    GSize := (ARect.Right - ARect.Left) - 1;

  ChunkSize := GSize div 3;
  if ChunkSize = 0 then ChunkSize := 1;
  d1 := ChunkSize;
  d2 := ChunkSize * 2;
  d3 := GSize;

  SpGradient(ACanvas, ARect, 0, d1, ChunkSize, C1, C2, Vertical);
  SpGradient(ACanvas, ARect, d1, d2, ChunkSize, C2, C3, Vertical);
  SpGradient(ACanvas, ARect, d2, d3, ChunkSize, C3, C4, Vertical);
end;

procedure SpGradientFill9pixels(ACanvas: TCanvas; const ARect: TRect;
  const C1, C2, C3, C4: TColor; const Vertical: Boolean);
// Mimics Vista menubar/toolbar blue gradient
var
  GSize, d1, d2: Integer;
begin
  if Vertical then
    GSize := (ARect.Bottom - ARect.Top) - 1
  else
    GSize := (ARect.Right - ARect.Left) - 1;

  d1 := GSize div 3;
  if d1 > 9 then d1 := 9;
  d2 := GSize;

  SpGradient(ACanvas, ARect, 0, d1, d1, C1, C2, Vertical);
  SpGradient(ACanvas, ARect, d1, d2, d2 - d1, C3, C4, Vertical);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Element painting }

procedure SpDrawArrow(ACanvas: TCanvas; X, Y: Integer; AColor: TColor; Vertical, Reverse: Boolean; Size: Integer);
var
  C1, C2: TColor;
begin
  C1 := ACanvas.Pen.Color;
  C2 := ACanvas.Brush.Color;
  ACanvas.Brush.Color := AColor;
  ACanvas.Pen.Color := AColor;
  ACanvas.Pen.Width := 1;

  if Vertical then
    if Reverse then
      ACanvas.Polygon([Point(X, Y), Point(X - Size, Y + Size), Point(X + Size, Y + Size)])
    else
      ACanvas.Polygon([Point(X - Size, Y), Point(X + Size, Y), Point(X, Y + Size)])
  else
    if Reverse then
      ACanvas.Polygon([Point(X, Y), Point(X + Size, Y + Size), Point(X + Size, Y - Size)])
    else
      ACanvas.Polygon([Point(X, Y - Size), Point(X, Y + Size), Point(X + Size, Y)]);

  ACanvas.Pen.Color := C1;
  ACanvas.Brush.Color := C2;
end;

procedure SpDrawDropMark(ACanvas: TCanvas; DropMark: TRect);
var
  C: TColor;
  R: TRect;
begin
  if IsRectEmpty(DropMark) then Exit;
  C := ACanvas.Brush.Color;

  R := Rect(DropMark.Left + 1, DropMark.Top, DropMark.Right - 1, DropMark.Top + 2);
  ACanvas.Rectangle(R);
  R := Rect(DropMark.Left + 1, DropMark.Bottom - 2, DropMark.Right - 1, DropMark.Bottom);
  ACanvas.Rectangle(R);

  R := Rect(DropMark.Left, DropMark.Top + 1, DropMark.Right, DropMark.Top + 3);
  ACanvas.Rectangle(R);
  R := Rect(DropMark.Left, DropMark.Bottom - 3, DropMark.Right, DropMark.Bottom - 1);
  ACanvas.Rectangle(R);

  R := Rect(DropMark.Left + 1, DropMark.Top + 4, DropMark.Right - 1, DropMark.Bottom - 4);
  ACanvas.Rectangle(R);

    {
    // Standard DropMark

    R := Rect(DropMark.Left, DropMark.Top, DropMark.Right - 1, DropMark.Bottom - 1);
    if IsRectEmpty(R) then Exit;
    ACanvas.Brush.Color := clBlack;
    ACanvas.Polygon([
      Point(R.Left, R.Top),
      Point(R.Left + 2, R.Top + 2),
      Point(R.Left + 2, R.Bottom - 2),
      Point(R.Left, R.Bottom),
      Point(R.Right, R.Bottom),
      Point(R.Right - 2, R.Bottom - 2),
      Point(R.Right - 2, R.Top + 2),
      Point(R.Right, R.Top)
    ]);
    }
  ACanvas.Brush.Color := C;
end;

procedure SpDrawFocusRect(ACanvas: TCanvas; const ARect: TRect);
var
  DC: HDC;
  C1, C2: TColor;
begin
  if not IsRectEmpty(ARect) then begin
    DC := ACanvas.Handle;
    C1 := SetTextColor(DC, clBlack);
    C2 := SetBkColor(DC, clWhite);
    ACanvas.DrawFocusRect(ARect);
    SetTextColor(DC, C1);
    SetBkColor(DC, C2);
  end;
end;

procedure SpDrawGlyphPattern(ACanvas: TCanvas; ARect: TRect;
  Pattern: TSpTBXGlyphPattern; PatternColor: TColor; DPI: Integer);
var
  Size: Integer;
  PenStyle: Cardinal;
  R: TRect;
  BrushInfo: TLogBrush;
  PColor, BColor: TColor;
begin
  {
  Close, 7x7   Maximize, 8x8   Minimize, 8x8   Restore, 9x9
  xx---xx      xxxxxxxx        --------        --xxxxxx-
  xxx-xxx      xxxxxxxx        --------        --xxxxxx-
  -xxxxx-      x------x        --------        --x----x-
  --xxx--      x------x        --------        xxxxxx-x-
  -xxxxx-      x------x        --------        xxxxxx-x-
  xxx-xxx      x------x        -xxxxxx-        x----xxx-
  xx---xx      x------x        -xxxxxx-        x----x---
               xxxxxxxx        --------        x----x---
                                               xxxxxx---

  Toolbar Close, 6x6   Chevron, 8x8    Vertical Chevron, 8x8
  xx--xx               xx--xx--        x---x---
  -xxxx-               -xx--xx-        xx-xx---
  --xx--               --xx--xx        -xxx----
  -xxxx-               -xx--xx-        --x-----
  xx--xx               xx--xx--        x---x---
  ------               --------        xx-xx---
                       --------        -xxx----
                       --------        --x-----

  Menu Checkmark, 7x7  Checkmark, 7x7  Menu Radiomark, 6x6
  ------x              ------x         --xx--
  -----xx              -----xx         -xxxx-
  x---xx-              x---xxx         xxxxxx
  xx-xx--              xx-xxx-         xxxxxx
  -xxx---              xxxxx--         -xxxx-
  --x----              -xxx---         --xx--
  -------              --x----
  }
  ACanvas.Pen.Width := SpPPIScale(1, DPI);
  PenStyle := PS_GEOMETRIC or PS_ENDCAP_SQUARE or PS_SOLID or PS_JOIN_MITER;
  case Pattern of
    gptClose: begin
         Size := SpPPIScale(7, DPI);
         // Round EndCap/Join
         PenStyle := PS_GEOMETRIC or PS_ENDCAP_ROUND or PS_SOLID or PS_JOIN_ROUND;
       end;
    gptMaximize: Size := SpPPIScale(8, DPI);
    gptMinimize: Size := SpPPIScale(8, DPI);
    gptRestore: Size := SpPPIScale(9, DPI);
    gptToolbarClose: begin
         Size := SpPPIScale(6, DPI);
         // Round EndCap/Join
         PenStyle := PS_GEOMETRIC or PS_ENDCAP_ROUND or PS_SOLID or PS_JOIN_ROUND;
       end;
    gptChevron: Size := SpPPIScale(8, DPI);
    gptVerticalChevron: Size := SpPPIScale(8, DPI);
    gptMenuCheckmark, gptCheckmark: begin
        Size := SpPPIScale(7, DPI);
        if DPI > 96 then
          inc(Size);
      end;
    gptMenuRadiomark: begin
         Size := SpPPIScale(7, DPI);
         // Round EndCap/Join
         PenStyle := PS_GEOMETRIC or PS_ENDCAP_ROUND or PS_SOLID or PS_JOIN_ROUND;
         ACanvas.Pen.Width := 1; // always 1 pixel
       end;
  else
    Size := SpPPIScale(8, DPI);
  end;

  // Create a pen with a Square endcap
  BrushInfo.lbStyle := BS_SOLID;
  BrushInfo.lbColor := ColorToRGB(PatternColor);
  BrushInfo.lbHatch := 0;
  ACanvas.Pen.Handle := ExtCreatePen(PenStyle, ACanvas.Pen.Width, BrushInfo, 0, nil);

  // Center the pattern on ARect
  R := SpCenterRect(ARect, Size-1, Size-1);

  // Save colors
  PColor := ACanvas.Pen.Color;
  BColor := ACanvas.Brush.Color;

  case Pattern of
    gptClose:
      begin
        ACanvas.Polyline([
          Point(R.Left, R.Top + 1),
          Point(R.Right - 1, R.Bottom)
        ]);
        ACanvas.Polyline([
          Point(R.Left, R.Top),
          Point(R.Right, R.Bottom)
        ]);
        ACanvas.Polyline([
          Point(R.Left + 1, R.Top),
          Point(R.Right, R.Bottom - 1)
        ]);

        ACanvas.Polyline([
          Point(R.Right, R.Top + 1),
          Point(R.Left + 1, R.Bottom)
        ]);
        ACanvas.Polyline([
          Point(R.Right, R.Top),
          Point(R.Left, R.Bottom)
        ]);
        ACanvas.Polyline([
          Point(R.Right - 1, R.Top),
          Point(R.Left, R.Bottom - 1)
        ]);
      end;
    gptMaximize:
      begin
        ACanvas.Polyline([
          Point(R.Left, R.Top + SpPPIScale(1, DPI)),
          Point(R.Right, R.Top + SpPPIScale(1, DPI))
        ]);
        ACanvas.Polyline([
          Point(R.Left, R.Top),
          Point(R.Right, R.Top),
          Point(R.Right, R.Bottom),
          Point(R.Left, R.Bottom),
          Point(R.Left, R.Top)
        ]);
      end;
    gptMinimize:
      begin
        ACanvas.Polyline([
          Point(R.Left + SpPPIScale(1, DPI), R.Bottom - SpPPIScale(2, DPI)),
          Point(R.Right - SpPPIScale(1, DPI), R.Bottom - SpPPIScale(2, DPI)),
          Point(R.Right - SpPPIScale(1, DPI), R.Bottom - SpPPIScale(1, DPI)),
          Point(R.Left + SpPPIScale(1, DPI), R.Bottom - SpPPIScale(1, DPI)),
          Point(R.Left + SpPPIScale(1, DPI), R.Bottom - SpPPIScale(2, DPI))
        ]);
      end;
    gptRestore:
      begin
        ACanvas.Polyline([
          Point(R.Left + SpPPIScale(2, DPI), R.Top + SpPPIScale(2, DPI)),
          Point(R.Left + SpPPIScale(2, DPI), R.Top),
          Point(R.Left + SpPPIScale(7, DPI), R.Top),
          Point(R.Left + SpPPIScale(7, DPI), R.Top + SpPPIScale(5, DPI)),
          Point(R.Left + SpPPIScale(6, DPI), R.Top + SpPPIScale(5, DPI))
        ]);
        ACanvas.Polyline([
          Point(R.Left + SpPPIScale(2, DPI), R.Top + SpPPIScale(1, DPI)),
          Point(R.Left + SpPPIScale(7, DPI), R.Top + SpPPIScale(1, DPI))
        ]);
        ACanvas.Polyline([
          Point(R.Left, R.Top + SpPPIScale(3, DPI)),
          Point(R.Left + SpPPIScale(5, DPI), R.Top + SpPPIScale(3, DPI)),
          Point(R.Left + SpPPIScale(5, DPI), R.Bottom),
          Point(R.Left, R.Bottom),
          Point(R.Left, R.Top + SpPPIScale(3, DPI))
        ]);
        ACanvas.Polyline([
          Point(R.Left, R.Top + SpPPIScale(4, DPI)),
          Point(R.Left + SpPPIScale(5, DPI), R.Top + SpPPIScale(4, DPI))
        ]);
      end;
    gptToolbarClose:
      begin
        ACanvas.Polyline([
          Point(R.Left, R.Top),
          Point(R.Right - 1, R.Bottom - 1)
        ]);
        ACanvas.Polyline([
          Point(R.Left + 1, R.Top),
          Point(R.Right, R.Bottom - 1)
        ]);
        ACanvas.Polyline([
          Point(R.Right, R.Top),
          Point(R.Left + 1, R.Bottom - 1)
        ]);
        ACanvas.Polyline([
          Point(R.Right - 1, R.Top),
          Point(R.Left, R.Bottom - 1)
        ]);
      end;
    gptChevron:
      begin
        ACanvas.Polyline([
          Point(R.Left, R.Top),
          Point(R.Left + SpPPIScale(2, DPI), R.Top + SpPPIScale(2, DPI)),
          Point(R.Left, R.Top + SpPPIScale(2, DPI) * 2)
        ]);
        ACanvas.Polyline([
          Point(R.Left + 1, R.Top),
          Point(R.Left + SpPPIScale(2, DPI) + 1, R.Top + SpPPIScale(2, DPI)),
          Point(R.Left + 1, R.Top + SpPPIScale(2, DPI) * 2)
        ]);
        ACanvas.Polyline([
          Point(R.Left + SpPPIScale(4, DPI), R.Top),
          Point(R.Left + SpPPIScale(6, DPI), R.Top + SpPPIScale(2, DPI)),
          Point(R.Left + SpPPIScale(4, DPI), R.Top + SpPPIScale(2, DPI) * 2)
        ]);
        ACanvas.Polyline([
          Point(R.Left + SpPPIScale(4, DPI) + 1, R.Top),
          Point(R.Left + SpPPIScale(6, DPI) + 1, R.Top + SpPPIScale(2, DPI)),
          Point(R.Left + SpPPIScale(4, DPI) + 1, R.Top + SpPPIScale(2, DPI) * 2)
        ]);
      end;
    gptVerticalChevron:
      begin
        ACanvas.Polyline([
          Point(R.Left, R.Top),
          Point(R.Left + SpPPIScale(2, DPI), R.Top + SpPPIScale(2, DPI)),
          Point(R.Left + SpPPIScale(2, DPI) * 2, R.Top)
        ]);
        ACanvas.Polyline([
          Point(R.Left, R.Top + 1),
          Point(R.Left + SpPPIScale(2, DPI), R.Top + SpPPIScale(2, DPI) + 1),
          Point(R.Left + SpPPIScale(2, DPI) * 2, R.Top + 1)
        ]);

        ACanvas.Polyline([
          Point(R.Left, R.Top + SpPPIScale(4, DPI)),
          Point(R.Left + SpPPIScale(2, DPI), R.Top + SpPPIScale(6, DPI)),
          Point(R.Left + SpPPIScale(2, DPI) * 2, R.Top + SpPPIScale(4, DPI))
        ]);
        ACanvas.Polyline([
          Point(R.Left, R.Top + SpPPIScale(4, DPI) + 1),
          Point(R.Left + SpPPIScale(2, DPI), R.Top + SpPPIScale(6, DPI) + 1),
          Point(R.Left + SpPPIScale(2, DPI) * 2, R.Top + SpPPIScale(4, DPI) + 1)
        ]);
      end;
    gptMenuCheckmark:
      begin
        ACanvas.Polyline([
          Point(R.Left, R.Top + SpPPIScale(2, DPI)),
          Point(R.Left + SpPPIScale(2, DPI), R.Top + SpPPIScale(2, DPI) * 2),
          Point(R.Right, R.Top)
        ]);
        ACanvas.Polyline([
          Point(R.Left, R.Top + SpPPIScale(2, DPI) + 1),
          Point(R.Left + SpPPIScale(2, DPI), R.Top + SpPPIScale(2, DPI) * 2 + 1),
          Point(R.Right, R.Top + 1)
        ]);
      end;
    gptCheckmark:
      begin
        ACanvas.Polyline([
          Point(R.Left, R.Top + SpPPIScale(2, DPI)),
          Point(R.Left + SpPPIScale(2, DPI), R.Top + SpPPIScale(2, DPI) * 2),
          Point(R.Right, R.Top)
        ]);
        ACanvas.Polyline([
          Point(R.Left, R.Top + SpPPIScale(2, DPI) + 1),
          Point(R.Left + SpPPIScale(2, DPI), R.Top + SpPPIScale(2, DPI) * 2 + 1),
          Point(R.Right, R.Top + 1)
        ]);
        ACanvas.Polyline([
          Point(R.Left, R.Top + SpPPIScale(2, DPI) + 2),
          Point(R.Left + SpPPIScale(2, DPI), R.Top + SpPPIScale(2, DPI) * 2 + 2),
          Point(R.Right, R.Top + 2)
        ]);
      end;
    gptMenuRadiomark:
      begin
        ACanvas.Brush.Color := PatternColor;
        ACanvas.Ellipse(R);
      end;
  end;

  ACanvas.Pen.Width := 1;
  ACanvas.Pen.Color := PColor;
  ACanvas.Brush.Color := BColor;
end;

procedure SpDrawXPButton(ACanvas: TCanvas; ARect: TRect; Enabled, Pushed,
  HotTrack, Checked, Focused, Defaulted: Boolean; DPI: Integer);
var
  C: TColor;
  State: TSpTBXSkinStatesType;
begin
  case SkinManager.GetSkinType of
    sknNone:
      begin
        C := ACanvas.Brush.Color;
        ACanvas.Brush.Color := clBtnFace;
        ACanvas.FillRect(ARect);
        if Defaulted or Focused then begin
          ACanvas.Brush.Color := clWindowFrame;
          ACanvas.FrameRect(ARect);
          InflateRect(ARect, -1, -1);  // Reduce the Rect for the focus rect
        end;
        if Pushed or Checked then begin
          ACanvas.Brush.Color := clBtnShadow;
          ACanvas.FrameRect(ARect);
        end
        else
          DrawFrameControl(ACanvas.Handle, ARect, DFC_BUTTON, DFCS_BUTTONPUSH);
        ACanvas.Brush.Color := C;
      end;
    sknWindows, sknDelphiStyle:
      CurrentSkin.PaintThemedElementBackground(ACanvas, ARect, skncButton, Enabled, Pushed, HotTrack, Checked, Focused, Defaulted, False, DPI);
    sknSkin:
      begin
        State := CurrentSkin.GetState(Enabled, Pushed, HotTrack, Checked);
        CurrentSkin.PaintBackground(ACanvas, ARect, skncButton, State, True, True);
      end;
  end;

  if Focused then begin
    InflateRect(ARect, -3, -3);
    SpDrawFocusRect(ACanvas, ARect);
  end;
end;

procedure SpDrawXPCheckBoxGlyph(ACanvas: TCanvas; ARect: TRect; Enabled: Boolean;
  State: TCheckBoxState; HotTrack, Pushed: Boolean; DPI: Integer);
var
  Flags: Cardinal;
  SknState: TSpTBXSkinStatesType;
  CheckColor: TColor;
begin
  Flags := 0;
  case SkinManager.GetSkinType of
    sknNone:
      begin
        case State of
          cbChecked: Flags := DFCS_BUTTONCHECK or DFCS_CHECKED;
          cbGrayed: Flags := DFCS_BUTTON3STATE or DFCS_CHECKED;
          cbUnChecked: Flags := DFCS_BUTTONCHECK;
        end;
        if not Enabled then
          Flags := Flags or DFCS_INACTIVE;
        if Pushed then
          Flags := Flags or DFCS_PUSHED;
        DrawFrameControl(ACanvas.Handle, ARect, DFC_BUTTON, Flags);
      end;
    sknWindows, sknDelphiStyle:
      CurrentSkin.PaintThemedElementBackground(ACanvas, ARect, skncCheckBox, Enabled, Pushed, HotTrack, State = cbChecked, False, False, State = cbGrayed, DPI);
    sknSkin:
      begin
        SknState := CurrentSkin.GetState(Enabled, Pushed, HotTrack, State in [cbChecked, cbGrayed]);
        CurrentSkin.PaintBackground(ACanvas, ARect, skncCheckBox, SknState, True, True);
        if State = cbChecked then begin
          CheckColor := CurrentSkin.GetTextColor(skncCheckBox, SknState);
          SpDrawGlyphPattern(ACanvas, ARect, gptCheckmark, CheckColor, DPI);
        end
        else
          if State = cbGrayed then begin
            InflateRect(ARect, -SpPPIScale(3, DPI), -SpPPIScale(3, DPI));
            CheckColor := CurrentSkin.Options(skncCheckBox, sknsChecked).Borders.Color1;
            SpFillRect(ACanvas, ARect, CheckColor);
          end;
      end;
  end;
end;

procedure SpDrawXPRadioButtonGlyph(ACanvas: TCanvas; ARect: TRect; Enabled: Boolean;
  Checked, HotTrack, Pushed: Boolean; DPI: Integer);
var
  Size, Flags: Integer;
  SknState: TSpTBXSkinStatesType;
begin
  case SkinManager.GetSkinType of
    sknNone:
      begin
        Flags := DFCS_BUTTONRADIO;
        if Checked then
          Flags := Flags or DFCS_CHECKED;
        if not Enabled then
          Flags := Flags or DFCS_INACTIVE;
        if Pushed then
          Flags := Flags or DFCS_PUSHED;
        DrawFrameControl(ACanvas.Handle, ARect, DFC_BUTTON, Flags);
      end;
    sknWindows, sknDelphiStyle:
      CurrentSkin.PaintThemedElementBackground(ACanvas, ARect, skncRadioButton, Enabled, Pushed, HotTrack, Checked, False, False, False, DPI);
    sknSkin:
      begin
        SknState := CurrentSkin.GetState(Enabled, Pushed, HotTrack, Checked);
        // Keep it simple make the radio 13x13
        ARect := SpCenterRect(ARect, SpPPIScale(13, DPI), SpPPIScale(13, DPI));

        // Background
        BeginPath(ACanvas.Handle);
        ACanvas.Ellipse(ARect);
        EndPath(ACanvas.Handle);
        SelectClipPath(ACanvas.Handle, RGN_COPY);
        CurrentSkin.PaintBackground(ACanvas, ARect, skncRadioButton, SknState, True, False);
        SelectClipPath(ACanvas.Handle, 0);
        SelectClipRgn(ACanvas.Handle, 0);
        // Frame
        ACanvas.Brush.Style := bsClear;
        ACanvas.Pen.Color := CurrentSkin.Options(skncRadioButton, SknState).Borders.Color1;;
        ACanvas.Ellipse(ARect);
        // Radio
        if Checked then begin
          ACanvas.Brush.Color := CurrentSkin.GetTextColor(skncRadioButton, SknState);
          ACanvas.Pen.Color := ACanvas.Brush.Color;
          Size := SpPPIScale(5, DPI);
          ACanvas.Ellipse(SpCenterRect(ARect, Size, Size));
        end;
      end;
  end;
end;

procedure SpDrawXPEditFrame(ACanvas: TCanvas; ARect: TRect; Enabled, HotTrack,
  ClipContent, AutoAdjust: Boolean; DPI: Integer);
  //  ClipContent: Boolean = False; AutoAdjust: Boolean = False;
var
  BorderR: TRect;
  State: TSpTBXSkinStatesType;
  Entry: TSpTBXSkinOptionEntry;
begin
  if ClipContent then begin
    BorderR := ARect;
    if HotTrack then
      InflateRect(BorderR, -1, -1)
    else
      InflateRect(BorderR, -2, -2);
    ExcludeClipRect(ACanvas.Handle, BorderR.Left, BorderR.Top, BorderR.Right, BorderR.Bottom);
  end;
  try
    case SkinManager.GetSkinType of
      sknNone:
        if HotTrack then
          SpDrawRectangle(ACanvas, ARect, 0, clBtnShadow, clBtnHighlight, clBtnFace, clBtnFace)
        else
          SpDrawRectangle(ACanvas, ARect, 0, clBtnFace, clBtnFace, clBtnFace, clBtnFace);
      sknWindows, sknDelphiStyle:
        begin
          CurrentSkin.PaintThemedElementBackground(ACanvas, ARect, skncEditFrame, Enabled, False, HotTrack, False, False, False, False, DPI);
        end;
      sknSkin:
        begin
          State := CurrentSkin.GetState(Enabled, False, HotTrack, False);
          // Try to adjust the borders if only the internal borders are specified,
          // used by some controls that need to paint the edit frames like
          // TSpTBXPanel (HotTrack=True), TSpTBXListBox, TSpTBXCheckListBox, etc
          if AutoAdjust then begin
            Entry := SkinManager.CurrentSkin.Options(skncEditFrame, State).Borders;
            if (Entry.Color1 = clNone) and (Entry.Color2 = clNone) and
              (Entry.Color3 <> clNone) and (Entry.Color4 <> clNone) then
            begin
              CurrentSkin.PaintBackground(ACanvas, ARect, skncEditFrame, State, True, False);
              SpDrawRectangle(ACanvas, ARect, Entry.SkinType, Entry.Color3, Entry.Color4);
              Exit;
            end;
          end;

          CurrentSkin.PaintBackground(ACanvas, ARect, skncEditFrame, State, True, True);
        end;
    end;
  finally
    if ClipContent then
      SelectClipRgn(ACanvas.Handle, 0);
  end;
end;

procedure SpDrawXPEditFrame(AWinControl: TWinControl; HotTracking,
  AutoAdjust, HideFrame: Boolean; DPI: Integer);
  //AutoAdjust: Boolean = False; HideFrame: Boolean = False;
var
  R: TRect;
  DC: HDC;
  ACanvas: TCanvas;
begin
  DC := GetWindowDC(AWinControl.Handle);
  try
    ACanvas := TCanvas.Create;
    try
      ACanvas.Handle := DC;
      GetWindowRect(AWinControl.Handle, R);
      OffsetRect(R, -R.Left, -R.Top);
      with R do
        ExcludeClipRect(DC, Left + 2, Top + 2, Right - 2, Bottom - 2); // Do not use SpDPIScale

      if HideFrame then begin
        ACanvas.Brush.Color := TControlAccess(AWinControl).Color;
        ACanvas.FillRect(R);
      end
      else begin
        // Don't use SpDrawParentBackground to paint the background it doesn't get
        // the correct WindowOrg in this particular case
//  Commented because of conflict when painting on Glass
//        PerformEraseBackground(AWinControl, ACanvas.Handle);
        SpDrawParentBackground(AWinControl, ACanvas.Handle, R);

        SpDrawXPEditFrame(ACanvas, R, AWinControl.Enabled, HotTracking, False, AutoAdjust, DPI);
      end;
    finally
      ACanvas.Handle := 0;
      ACanvas.Free;
    end;
  finally
    ReleaseDC(AWinControl.Handle, DC);
  end;
end;

procedure SpDrawXPGrip(ACanvas: TCanvas; ARect: TRect; LoC, HiC: TColor; DPI: Integer);
var
  I, J: Integer;
  XCellCount, YCellCount: Integer;
  R: TRect;
  C: TColor;
begin
  //  4 x 4 cells (Grey, White, Null)
  //  GG--
  //  GGW-
  //  -WW-
  //  ----

  C := ACanvas.Brush.Color;
  XCellCount := (ARect.Right - ARect.Left) div SpPPIScale(4, DPI);
  YCellCount := (ARect.Bottom - ARect.Top) div SpPPIScale(4, DPI);
  if XCellCount = 0 then XCellCount := 1;
  if YCellCount = 0 then YCellCount := 1;

  for J := 0 to YCellCount - 1 do
    for I := 0 to XCellCount - 1 do begin
      R.Left := ARect.Left + (I * SpPPIScale(4, DPI)) + SpPPIScale(1, DPI);
      R.Right := R.Left + SpPPIScale(2, DPI);
      R.Top := ARect.Top + (J * SpPPIScale(4, DPI)) + SpPPIScale(1, DPI);
      R.Bottom := R.Top + SpPPIScale(2, DPI);

      ACanvas.Brush.Color := HiC;
      ACanvas.FillRect(R);
      OffsetRect(R, -SpPPIScale(1, DPI), -SpPPIScale(1, DPI));
      ACanvas.Brush.Color := LoC;
      ACanvas.FillRect(R);
    end;
  ACanvas.Brush.Color := C;
end;

procedure SpDrawXPHeader(ACanvas: TCanvas; ARect: TRect; HotTrack, Pushed: Boolean; DPI: Integer);
var
  State: TSpTBXSkinStatesType;
begin
  case SkinManager.GetSkinType of
    sknNone:
      begin
        DrawEdge(ACanvas.Handle, ARect, BDR_RAISEDINNER or BDR_RAISEDOUTER, BF_RECT or BF_SOFT);
      end;
    sknWindows, sknDelphiStyle:
      begin
        CurrentSkin.PaintThemedElementBackground(ACanvas, ARect, skncHeader, True, Pushed, HotTrack, False, False, False, False, DPI);
      end;
    sknSkin:
      begin
        State := CurrentSkin.GetState(True, Pushed, HotTrack, False);
        if (State = sknsPushed) and CurrentSkin.Options(skncHeader, State).IsEmpty then
          State := sknsHotTrack;
        CurrentSkin.PaintBackground(ACanvas, ARect, skncHeader, State, True, True);
      end;
  end;
end;

procedure SpDrawXPListItemBackground(ACanvas: TCanvas; ARect: TRect; Selected, HotTrack, Focused: Boolean;
  ForceRectBorders: Boolean; Borders: Boolean);
var
  State: TSpTBXSkinStatesType;
begin
  State := CurrentSkin.GetState(True, False, HotTrack, Selected);
  ACanvas.Font.Color := CurrentSkin.GetTextColor(skncListItem, State);

  if SkinManager.GetSkinType = sknSkin then begin
    ACanvas.FillRect(ARect);
    if HotTrack or Selected then begin
      if ForceRectBorders then
        CurrentSkin.PaintBackground(ACanvas, ARect, skncListItem, State, True, Borders, False, [akLeft, akTop, akRight, akBottom])
      else
        CurrentSkin.PaintBackground(ACanvas, ARect, skncListItem, State, True, Borders);
    end;
  end
  else begin
    if SkinManager.GetSkinType = sknDelphiStyle then begin
      {$IF CompilerVersion >= 23} // for Delphi XE2 and up
      if Selected then
        ACanvas.Brush.Color := SpTBXThemeServices.GetSystemColor(clHighlight)
      else
        ACanvas.Brush.Color := SpTBXThemeServices.GetStyleColor(scListBox);
      {$IFEND}
    end
    else
      if Selected then
        ACanvas.Brush.Color := clHighlight;
    ACanvas.FillRect(ARect);
    if Focused then
      SpDrawFocusRect(ACanvas, ARect);
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Skins painting }

procedure SpPaintSkinBackground(ACanvas: TCanvas; ARect: TRect; SkinOption: TSpTBXSkinOptionCategory; Vertical: Boolean);
var
  Part: TSpTBXSkinOptionEntry;
  SkinType: Integer;
begin
  Part := SkinOption.Body;
  SkinType := SkinOption.Body.SkinType;

  if Vertical then
    case SkinType of
      1: SkinType := 2;  // Vertical Gradient to Horizontal
      2: SkinType := 1;  // Horizontal Gradient to Vertical

      3: SkinType := 4;  // Vertical Glass Gradient to Horizontal
      4: SkinType := 3;  // Horizontal Glass Gradient to Vertical

      5: SkinType := 6;  // Vertical Mirror Gradient to Horizontal
      6: SkinType := 5;  // Horizontal Mirror Gradient to Vertical

      7: SkinType := 8;  // Vertical MirrorTop Gradient to Horizontal
      8: SkinType := 7;  // Horizontal MirrorTop Gradient to Vertical

      9: SkinType := 10; // Vertical 9Pixels Gradient to Horizontal
      10: SkinType := 9; // Horizontal 9Pixels Gradient to Vertical
    end;

  case SkinType of
    0: begin  // Solid
         SpFillRect(ACanvas, ARect, Part.Color1);
       end;
    1: begin  // Vertical Gradient
         SpGradientFill(ACanvas, ARect, Part.Color1, Part.Color2, True);
       end;
    2: begin  // Horizontal Gradient
         SpGradientFill(ACanvas, ARect, Part.Color1, Part.Color2, False);
       end;
    3: begin  // Vertical Glass Gradient
         SpGradientFillGlass(ACanvas, ARect, Part.Color1, Part.Color2, Part.Color3, Part.Color4, True);
       end;
    4: begin  // Horizontal Glass Gradient
         SpGradientFillGlass(ACanvas, ARect, Part.Color1, Part.Color2, Part.Color3, Part.Color4, False);
       end;
    5: begin  // Vertical Mirror Gradient
         SpGradientFillMirror(ACanvas, ARect, Part.Color1, Part.Color2, Part.Color3, Part.Color4, True);
       end;
    6: begin  // Horizontal Mirror Gradient
         SpGradientFillMirror(ACanvas, ARect, Part.Color1, Part.Color2, Part.Color3, Part.Color4, False);
       end;
    7: begin  // Vertical MirrorTop Gradient
         SpGradientFillMirrorTop(ACanvas, ARect, Part.Color1, Part.Color2, Part.Color3, Part.Color4, True);
       end;
    8: begin  // Horizontal MirrorTop Gradient
         SpGradientFillMirrorTop(ACanvas, ARect, Part.Color1, Part.Color2, Part.Color3, Part.Color4, False);
       end;
    9: begin  // Vertical 9Pixels Gradient
         SpGradientFill9Pixels(ACanvas, ARect, Part.Color1, Part.Color2, Part.Color3, Part.Color4, True);
       end;
    10:begin  // Horizontal 9Pixels Gradient
         SpGradientFill9Pixels(ACanvas, ARect, Part.Color1, Part.Color2, Part.Color3, Part.Color4, False);
       end;
  end;
end;

procedure SpPaintSkinBorders(ACanvas: TCanvas; ARect: TRect; SkinOption: TSpTBXSkinOptionCategory;
  ForceRectBorders: TAnchors = []);
var
  Part: TSpTBXSkinOptionEntry;
begin
  Part := SkinOption.Borders;
  case Part.SkinType of
    0, 1, 2: // Rectangle, Simple Rounded and Double Rounded Border
      begin
        SpDrawRectangle(ACanvas, ARect, Part.SkinType, Part.Color1, Part.Color2, Part.Color3, Part.Color4, ForceRectBorders);
      end;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Misc }

function SpIsWinVistaOrUp: Boolean;
begin
  Result := (Win32Platform = VER_PLATFORM_WIN32_NT) and (Win32MajorVersion >= 6);
end;

function SpIsWin10OrUp: Boolean;
begin
  Result := (Win32Platform = VER_PLATFORM_WIN32_NT) and (Win32MajorVersion >= 10);
end;

function SpGetDirectories(Path: string; L: TStringList): Boolean;
var
  SearchRec: TSearchRec;
begin
  Result := False;
  if DirectoryExists(Path) then begin
    Path := IncludeTrailingPathDelimiter(Path) + '*.*';
    if FindFirst(Path, faDirectory, SearchRec) = 0 then begin
      try
        repeat
          if (SearchRec.Name <> '.') and (SearchRec.Name <> '..') then
            L.Add(SearchRec.Name);
        until FindNext(SearchRec) <> 0;
        Result := True;
      finally
        FindClose(SearchRec);
      end;
    end;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ DPI }

function SpPPIScale(Value, DPI: Integer): Integer;
begin
  if DPI <= 0 then DPI := 96;
  Result := MulDiv(Value, DPI, 96);
end;

function SpPPIScaleToDPI(PPIScale: TPPIScale): Integer;
begin
  Result := MulDiv(PPIScale(96), 96, 96);
end;

procedure SpDPIResizeBitmap(Bitmap: TBitmap; const NewWidth, NewHeight, DPI: Integer);
var
  B: TBitmap;
begin
  B := TBitmap.Create;
  try
    B.SetSize(NewWidth, NewHeight);
    if DPI * 100 / 96 >= 150 then begin // Stretch if >= 150%
      SetStretchBltMode(B.Canvas.Handle, STRETCH_HALFTONE);
      B.Canvas.StretchDraw(Rect(0, 0, NewWidth, NewHeight), Bitmap);
    end
    else // Center if < 150%
      B.Canvas.Draw((NewWidth - Bitmap.Width) div 2, (NewHeight - Bitmap.Height) div 2, Bitmap);

    Bitmap.SetSize(NewWidth, NewHeight);
    Bitmap.Canvas.Draw(0, 0, B);
  finally
    B.Free;
  end;
end;

// newpy check with newer version of SpDPIScaleImageList
{
procedure SpDPIScaleImageList(const ImageList: TCustomImageList; M, D: Integer);
const
  ANDbits: array[0..2*16-1] of  Byte = ($FF,$FF,$FF,$FF,$FF,$FF,$FF,$FF,
                                        $FF,$FF,$FF,$FF,$FF,$FF,$FF,$FF,
                                        $FF,$FF,$FF,$FF,$FF,$FF,$FF,$FF,
                                        $FF,$FF,$FF,$FF,$FF,$FF,$FF,$FF);
  XORbits: array[0..2*16-1] of  Byte = ($00,$00,$00,$00,$00,$00,$00,$00,
                                        $00,$00,$00,$00,$00,$00,$00,$00,
                                        $00,$00,$00,$00,$00,$00,$00,$00,
                                        $00,$00,$00,$00,$00,$00,$00,$00);
var
  I: integer;
  Icon: HIcon;
  TempIL : TCustomImageList;
begin
  if M = D then
    Exit;
  TempIL := TCustomImageList.CreateSize(MulDiv(ImageList.Width, M, D), MulDiv(ImageList.Height, M, D));
  try
    TempIL.ColorDepth := cd32Bit;
    TempIL.DrawingStyle := ImageList.DrawingStyle;
    TempIL.BkColor := ImageList.BkColor;
    TempIL.BlendColor := ImageList.BlendColor;
    for I := 0 to ImageList.Count-1 do
    begin
      Icon := ImageList_GetIcon(ImageList.Handle, I, LR_DEFAULTCOLOR);
      if Icon = 0 then
      begin
        Icon := CreateIcon(hInstance,16,16,1,1,@ANDbits,@XORbits);
      end;
      ImageList_AddIcon(TempIL.Handle, Icon);
      DestroyIcon(Icon);
    end;
    ImageList.Assign(TempIL);
  finally
    TempIL.Free;
  end;
end;

}
procedure SpDPIScaleImageList(const ImageList: TCustomImageList; M, D: Integer);
var
  I: integer;
  Bimage, Bmask: TBitmap;
  TempIL : TImageList;
begin
   if M = D then Exit;

  TempIL := TImageList.Create(nil);
  try
    // Set size to match DPI (like 250% of 16px = 40px)
    TempIL.Assign(ImageList);
    ImageList.Clear;
    ImageList.SetSize(MulDiv(ImageList.Width, M, D), MulDiv(ImageList.Height, M, D));

    // Add images back to original ImageList
    for I := 0 to -1 + TempIL.Count do begin
      Bimage := TBitmap.Create;
      Bmask := TBitmap.Create;
      try
        // Get the image bitmap
        Bimage.SetSize(TempIL.Width, TempIL.Height);
        Bimage.Canvas.FillRect(Bimage.Canvas.ClipRect);
        ImageList_DrawEx(TempIL.Handle, I, Bimage.Canvas.Handle, 0, 0, Bimage.Width, Bimage.Height, CLR_NONE, CLR_NONE, ILD_NORMAL);
        SpDPIResizeBitmap(Bimage, ImageList.Width, ImageList.Height, M); // Resize

        // Get the mask bitmap
        Bmask.SetSize(TempIL.Width, TempIL.Height);
        Bmask.Canvas.FillRect(Bmask.Canvas.ClipRect);
        ImageList_DrawEx(TempIL.Handle, I, Bmask.Canvas.Handle, 0, 0, Bmask.Width, Bmask.Height, CLR_NONE, CLR_NONE, ILD_MASK);
        SpDPIResizeBitmap(Bmask, ImageList.Width, ImageList.Height, M); // Resize

        // Add the bitmaps
        ImageList.Add(Bimage, Bmask);
      finally
        Bimage.Free;
        Bmask.Free;
      end;
    end;
  finally
    TempIL.Free;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXSkinOptionEntry }

procedure TSpTBXSkinOptionEntry.AssignTo(Dest: TPersistent);
begin
  if Dest is TSpTBXSkinOptionEntry then
    with TSpTBXSkinOptionEntry(Dest) do begin
      SkinType := Self.SkinType;
      Color1 := Self.Color1;
      Color2 := Self.Color2;
      Color3 := Self.Color3;
      Color4 := Self.Color4;
    end
  else inherited AssignTo(Dest);
end;

constructor TSpTBXSkinOptionEntry.Create;
begin
  inherited;
  Reset;
end;

procedure TSpTBXSkinOptionEntry.Fill(ASkinType: Integer; AColor1, AColor2,
  AColor3, AColor4: TColor);
begin
  FSkinType := ASkinType;
  FColor1 := AColor1;
  FColor2 := AColor2;
  FColor3 := AColor3;
  FColor4 := AColor4;
end;

function TSpTBXSkinOptionEntry.IsEmpty: Boolean;
begin
  Result := (FColor1 = clNone) and (FColor2 = clNone) and (FColor3 = clNone) and (FColor4 = clNone);
end;

function TSpTBXSkinOptionEntry.IsEqual(AOptionEntry: TSpTBXSkinOptionEntry): Boolean;
begin
  Result := (FSkinType = AOptionEntry.SkinType) and
    (FColor1 = AOptionEntry.Color1) and (FColor2 = AOptionEntry.Color2) and
    (FColor3 = AOptionEntry.Color3) and (FColor4 = AOptionEntry.Color4);
end;

procedure TSpTBXSkinOptionEntry.Lighten(Amount: Integer);
begin
  if FColor1 <> clNone then FColor1 := SpLighten(FColor1, Amount);
  if FColor2 <> clNone then FColor2 := SpLighten(FColor2, Amount);
  if FColor3 <> clNone then FColor3 := SpLighten(FColor3, Amount);
  if FColor4 <> clNone then FColor4 := SpLighten(FColor4, Amount);
end;

procedure TSpTBXSkinOptionEntry.Reset;
begin
  FSkinType := 0;
  FColor1 := clNone;
  FColor2 := clNone;
  FColor3 := clNone;
  FColor4 := clNone;
end;

procedure TSpTBXSkinOptionEntry.ReadFromString(S: string);
var
  L: TStringList;
begin
  Reset;
  L := TStringList.Create;
  try
    L.CommaText := S;
    try
      if L.Count > 0 then FSkinType := StrToIntDef(L[0], 0);
      if L.Count > 1 then FColor1 := StringToColor(L[1]);
      if L.Count > 2 then FColor2 := StringToColor(L[2]);
      if L.Count > 3 then FColor3 := StringToColor(L[3]);
      if L.Count > 4 then FColor4 := StringToColor(L[4]);
    except
      // do nothing
    end;
  finally
    L.Free;
  end;
end;

function TSpTBXSkinOptionEntry.WriteToString: string;
begin
  Result := Format('%d, %s, %s, %s, %s', [FSkinType,
    SpColorToString(FColor1), SpColorToString(FColor2),
    SpColorToString(FColor3), SpColorToString(FColor4)]);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXThemeOptionCategory }

procedure TSpTBXSkinOptionCategory.AssignTo(Dest: TPersistent);
begin
  if Dest is TSpTBXSkinOptionCategory then
    with TSpTBXSkinOptionCategory(Dest) do begin
      Body.Assign(Self.Body);
      Borders.Assign(Self.Borders);
      TextColor := Self.TextColor;
    end
  else inherited AssignTo(Dest);
end;

constructor TSpTBXSkinOptionCategory.Create;
begin
  inherited;
  FBody := TSpTBXSkinOptionEntry.Create;
  FBorders := TSpTBXSkinOptionEntry.Create;
  FTextColor := clNone;
end;

destructor TSpTBXSkinOptionCategory.Destroy;
begin
  FreeAndNil(FBody);
  FreeAndNil(FBorders);
  inherited;
end;

function TSpTBXSkinOptionCategory.IsEmpty: Boolean;
begin
  Result := FBody.IsEmpty and FBorders.IsEmpty and (FTextColor = clNone);
end;

procedure TSpTBXSkinOptionCategory.Reset;
begin
  FBody.Reset;
  FBorders.Reset;
  FTextColor := clNone;
end;

procedure TSpTBXSkinOptionCategory.SaveToIni(MemIni: TMemIniFile; Section, Ident: string);
begin
  MemIni.WriteString(Section, Ident + '.Body', Body.WriteToString);
  MemIni.WriteString(Section, Ident + '.Borders', Borders.WriteToString);
  MemIni.WriteString(Section, Ident + '.TextColor', SpColorToString(TextColor));
end;

procedure TSpTBXSkinOptionCategory.LoadFromIni(MemIni: TMemIniFile; Section, Ident: string);
begin
  Reset;
  if Ident = '' then Ident := SSpTBXSkinStatesString[sknsNormal];

  Body.ReadFromString(MemIni.ReadString(Section, Ident + '.Body', ''));
  Borders.ReadFromString(MemIni.ReadString(Section, Ident + '.Borders', ''));
  TextColor := StringToColor(MemIni.ReadString(Section, Ident + '.TextColor', 'clNone'));
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXThemeOptions }

procedure TSpTBXSkinOptions.AssignTo(Dest: TPersistent);
var
  C: TSpTBXSkinComponentsType;
  S: TSpTBXSkinStatesType;
  DestOp: TSpTBXSkinOptions;
begin
  if Dest is TSpTBXSkinOptions then begin
    DestOp := TSpTBXSkinOptions(Dest);
    for C := Low(C) to High(C) do
      for S := Low(S) to High(S) do
        DestOp.FOptions[C, S].Assign(Options(C, S));
    DestOp.ColorBtnFace := FColorBtnFace;
    DestOp.FloatingWindowBorderSize := FFloatingWindowBorderSize;
    DestOp.OfficeIcons := FOfficeIcons;
    DestOp.OfficeMenu := FOfficeMenu;
    DestOp.OfficeStatusBar := FOfficeStatusBar;
    DestOp.FSkinAuthor := FSkinAuthor;
    DestOp.FSkinName := FSkinName;
  end
  else inherited AssignTo(Dest);
end;

procedure TSpTBXSkinOptions.BroadcastChanges;
begin
  if Self = SkinManager.CurrentSkin then begin
    SkinManager.ResetToSystemStyle;
    SkinManager.BroadcastSkinNotification;
  end;
end;

procedure TSpTBXSkinOptions.CopyOptions(AComponent, ToComponent: TSpTBXSkinComponentsType);
var
  S: TSpTBXSkinStatesType;
begin
  for S := Low(S) to High(S) do
    FOptions[AComponent, S].AssignTo(FOptions[ToComponent, S]);
end;

constructor TSpTBXSkinOptions.Create;
var
  C: TSpTBXSkinComponentsType;
  S: TSpTBXSkinStatesType;
begin
  inherited;
  FSkinName := 'Default';
  FColorBtnFace := clBtnFace;
  FFloatingWindowBorderSize := 4;
  for C := Low(C) to High(C) do
    for S := Low(S) to High(S) do
      FOptions[C, S] := TSpTBXSkinOptionCategory.Create;

  FillOptions;
end;

destructor TSpTBXSkinOptions.Destroy;
var
  C: TSpTBXSkinComponentsType;
  S: TSpTBXSkinStatesType;
begin
  for C := Low(C) to High(C) do
    for S := Low(S) to High(S) do
      FreeAndNil(FOptions[C, S]);

  inherited;
end;

procedure TSpTBXSkinOptions.FillOptions;
begin
  // Used by descendants to fill the skin options
end;

procedure TSpTBXSkinOptions.Reset(ForceResetSkinProperties: Boolean);
var
  C: TSpTBXSkinComponentsType;
  S: TSpTBXSkinStatesType;
begin
  if ForceResetSkinProperties then begin
    FColorBtnFace := clBtnFace;
    FFloatingWindowBorderSize := 4;
    FOfficeIcons := False;
    FOfficeMenu := False;
    FOfficeStatusBar := False;
  end;

  for C := Low(C) to High(C) do
    for S := Low(S) to High(S) do
      FOptions[C, S].Reset;
end;

function TSpTBXSkinOptions.Options(Component: TSpTBXSkinComponentsType; State: TSpTBXSkinStatesType): TSpTBXSkinOptionCategory;
begin
  if CSpTBXSkinComponents[Component].States = [sknsNormal] then
    State := sknsNormal;
  Result := FOptions[Component, State];
end;

function TSpTBXSkinOptions.Options(Component: TSpTBXSkinComponentsType): TSpTBXSkinOptionCategory;
begin
  Result := FOptions[Component, sknsNormal];
end;

procedure TSpTBXSkinOptions.SaveToFile(Filename: string);
var
  MemIni: TMemIniFile;
begin
  MemIni := TMemIniFile.Create(Filename);
  try
    SaveToMemIni(MemIni);
    MemIni.UpdateFile;
  finally
    MemIni.Free;
  end;
end;

procedure TSpTBXSkinOptions.SaveToMemIni(MemIni: TMemIniFile);
var
  C: TSpTBXSkinComponentsType;
  S: TSpTBXSkinStatesType;
begin
  MemIni.WriteString('Skin', 'Name', FSkinName);
  MemIni.WriteString('Skin', 'Author', FSkinAuthor);
  MemIni.WriteString('Skin', 'ColorBtnFace', SpColorToString(FColorBtnFace));
  MemIni.WriteInteger('Skin', 'FloatingWindowBorderSize', FFloatingWindowBorderSize);
  MemIni.WriteBool('Skin', 'OfficeIcons', FOfficeIcons);
  MemIni.WriteBool('Skin', 'OfficeMenu', FOfficeMenu);
  MemIni.WriteBool('Skin', 'OfficeStatusBar', FOfficeStatusBar);

  for C := Low(C) to High(C) do begin
    for S := Low(S) to High(S) do
      if S in CSpTBXSkinComponents[C].States then
        FOptions[C, S].SaveToIni(MemIni, CSpTBXSkinComponents[C].Name, SSpTBXSkinStatesString[S]);
  end;
end;

procedure TSpTBXSkinOptions.SaveToStrings(L: TStrings);
var
  MemIni: TMemIniFile;
begin
  MemIni := TMemIniFile.Create('');
  try
    MemIni.SetStrings(L); // Transfer L contents to MemIni
    SaveToMemIni(MemIni);
    L.Clear;
    MemIni.GetStrings(L); // Transfer MemIni contents to L
  finally
    MemIni.Free;
  end;
end;

procedure TSpTBXSkinOptions.LoadFromFile(Filename: string);
var
  L: TStringList;
begin
  if FileExists(Filename) then begin
    L := TStringList.Create;
    try
      L.LoadFromFile(Filename);
      LoadFromStrings(L);
    finally
      L.Free;
    end;
  end;
end;

procedure TSpTBXSkinOptions.LoadFromStrings(L: TStrings);
var
  MemIni: TMemIniFile;
  C: TSpTBXSkinComponentsType;
  S: TSpTBXSkinStatesType;
begin
  MemIni := TMemIniFile.Create('');
  try
    MemIni.SetStrings(L);

    FSkinName := MemIni.ReadString('Skin', 'Name', '');
    FSkinAuthor := MemIni.ReadString('Skin', 'Author', '');
    FColorBtnFace := StringToColor(MemIni.ReadString('Skin', 'ColorBtnFace', 'clBtnFace'));
    FFloatingWindowBorderSize := MemIni.ReadInteger('Skin', 'FloatingWindowBorderSize', 4);
    FOfficeIcons := MemIni.ReadBool('Skin', 'OfficeIcons', False);
    FOfficeMenu := MemIni.ReadBool('Skin', 'OfficeMenu', False);
    FOfficeStatusBar := MemIni.ReadBool('Skin', 'OfficeStautsBar', False);

    for C := Low(C) to High(C) do begin
      for S := Low(S) to High(S) do
        if S in CSpTBXSkinComponents[C].States then
          FOptions[C, S].LoadFromIni(MemIni, CSpTBXSkinComponents[C].Name, SSpTBXSkinStatesString[S]);
    end;

    BroadcastChanges;
  finally
    MemIni.Free;
  end;
end;

function TSpTBXSkinOptions.GetOfficeIcons: Boolean;
// OfficeIcons is used to paint the menu items icons with Office XP shadows.
begin
  Result := FOfficeIcons and (SkinManager.GetSkinType = sknSkin);
end;

function TSpTBXSkinOptions.GetOfficeMenu: Boolean;
// When OfficeMenu is True the height of the separators on popup menus
// is 6 pixels, otherwise the size is 10 pixels.
// And when the item is disabled the hottrack is not painted.
begin
  Result := FOfficeMenu and (SkinManager.GetSkinType = sknSkin);
end;

function TSpTBXSkinOptions.GetOfficePopup: Boolean;
// OfficePopup is used to paint the PopupWindow with Office XP style.
// It is also used to paint the opened toolbar item with shadows.
begin
  Result := (SkinManager.GetSkinType = sknSkin) and not Options(skncOpenToolbarItem).IsEmpty;
end;

function TSpTBXSkinOptions.GetOfficeStatusBar: Boolean;
// OfficeStatusBar is used to paint the StatusBar panels with Office XP style.
var
  T: TSpTBXSkinType;
begin
  T := SkinManager.GetSkinType;
  Result := (FOfficeStatusBar and (T = sknSkin)) or (T = sknNone);
end;

function TSpTBXSkinOptions.GetFloatingWindowBorderSize: Integer;
begin
  if SkinManager.GetSkinType = sknSkin then
    Result := FFloatingWindowBorderSize
  else
    Result := 4;
end;

procedure TSpTBXSkinOptions.SetFloatingWindowBorderSize(const Value: Integer);
begin
  FFloatingWindowBorderSize := Value;
  if FFloatingWindowBorderSize < 0 then FFloatingWindowBorderSize := 0;
  if FFloatingWindowBorderSize > 4 then FFloatingWindowBorderSize := 4;
end;

procedure TSpTBXSkinOptions.GetMenuItemMargins(ACanvas: TCanvas; ImgSize: Integer;
  out MarginsInfo: TSpTBXMenuItemMarginsInfo; DPI: Integer);
var
  TextMetric: TTextMetric;
  H, M2: Integer;
  SkinType: TSpTBXSkinType;
begin
  if ImgSize = 0 then
    ImgSize := SpPPIScale(16, DPI);

  FillChar(MarginsInfo, SizeOf(MarginsInfo), 0);
  SkinType := SkinManager.GetSkinType;

  if ((SkinType = sknWindows) and not SpIsWinVistaOrUp) or (SkinType = sknNone) then begin
    MarginsInfo.Margins := Rect(0, SpPPIScale(2, DPI), 0, SpPPIScale(2, DPI)); // MID_MENUITEM
    MarginsInfo.ImageTextSpace := SpPPIScale(1, DPI);         // TMI_MENU_IMGTEXTSPACE
    MarginsInfo.LeftCaptionMargin := SpPPIScale(2, DPI);      // TMI_MENU_LCAPTIONMARGIN
    MarginsInfo.RightCaptionMargin := SpPPIScale(2, DPI);     // TMI_MENU_RCAPTIONMARGIN
  end
  else begin
    // Vista-like spacing
    MarginsInfo.Margins := Rect(SpPPIScale(1, DPI), SpPPIScale(3, DPI), SpPPIScale(1, DPI), SpPPIScale(3, DPI)); // MID_MENUITEM
    MarginsInfo.ImageTextSpace := SpPPIScale(5 + 1, DPI);     // TMI_MENU_IMGTEXTSPACE
    MarginsInfo.LeftCaptionMargin := SpPPIScale(3, DPI);      // TMI_MENU_LCAPTIONMARGIN
    MarginsInfo.RightCaptionMargin := SpPPIScale(3, DPI);     // TMI_MENU_RCAPTIONMARGIN
  end;

  GetTextMetrics(ACanvas.Handle, TextMetric);
  M2 := MarginsInfo.Margins.Top + MarginsInfo.Margins.Bottom;
  MarginsInfo.GutterSize := TextMetric.tmHeight + TextMetric.tmExternalLeading + M2;
  H := ImgSize + M2;
  if H > MarginsInfo.GutterSize then MarginsInfo.GutterSize := H;
  MarginsInfo.GutterSize := (ImgSize + M2) * MarginsInfo.GutterSize div H;  // GutterSize = GetPopupMargin = ItemInfo.PopupMargin
end;

function TSpTBXSkinOptions.GetState(Enabled, Pushed, HotTrack, Checked: Boolean): TSpTBXSkinStatesType;
begin
  Result := sknsNormal;
  if not Enabled then Result := sknsDisabled
  else begin
    if Pushed then Result := sknsPushed
    else
      if HotTrack and Checked then Result := sknsCheckedAndHotTrack
      else
        if HotTrack then Result := sknsHotTrack
        else
          if Checked then Result := sknsChecked;
  end;
end;

procedure TSpTBXSkinOptions.GetState(State: TSpTBXSkinStatesType; out Enabled,
  Pushed, HotTrack, Checked: Boolean);
begin
  Enabled := State <> sknsDisabled;
  Pushed := State = sknsPushed;
  HotTrack := (State = sknsCheckedAndHotTrack) or (State = sknsHotTrack);
  Checked := (State = sknsCheckedAndHotTrack) or (State = sknsChecked);
end;

function TSpTBXSkinOptions.GetTextColor(Component: TSpTBXSkinComponentsType;
  State: TSpTBXSkinStatesType): TColor;
var
  Details: TThemedElementDetails;
  SkinType: TSpTBXSkinType;
begin
  Result := clNone;
  SkinType := SkinManager.GetSkinType;

  if SkinType = sknSkin then begin
    if State in CSpTBXSkinComponents[Component].States then begin
      Result := Options(Component, State).TextColor;
      if Result <> clNone then
        Exit; // Text color is specified by the skin
    end
    else
      Exit; // Exit if the State is not valid
  end;

 if State = sknsDisabled then Result := clGrayText
 else Result := clBtnText;
 if SkinType = sknDelphiStyle then
   Result := GetThemedSystemColor(Result);

  case Component of
    skncMenuItem:
      if ((SkinType = sknWindows) and SpIsWinVistaOrUp) or (SkinType = sknDelphiStyle) then begin
        // Use the new API on Windows Vista
        if GetThemedElementDetails(Component, State, Details) then
          GetThemedElementTextColor(Details, Result);
      end
      else
        if State <> sknsDisabled then begin
          Result := clMenuText;
          if SkinType <> sknSkin then
            if State in [sknsHotTrack, sknsCheckedAndHotTrack, sknsPushed] then
              Result := clHighlightText;
        end;
    skncMenuBarItem:
      if ((SkinType = sknWindows) and SpIsWinVistaOrUp) or (SkinType = sknDelphiStyle) then begin
        // Use the new API on Windows Vista
        if GetThemedElementDetails(Component, State, Details) then
          GetThemedElementTextColor(Details, Result);
      end
      else
        if State <> sknsDisabled then begin
          Result := clMenuText;
          if SkinType = sknWindows then
            if State in [sknsHotTrack, sknsPushed, sknsChecked, sknsCheckedAndHotTrack] then
              Result := clHighlightText;
        end;
    skncToolbarItem:
      if SkinType = sknDelphiStyle then begin
        if GetThemedElementDetails(Component, State, Details) then
          GetThemedElementTextColor(Details, Result);
      end
      else
        if State <> sknsDisabled then Result := clMenuText;
    skncButton, skncCheckBox, skncRadioButton, skncLabel, skncTab, skncEditButton:
      if (SkinType = sknWindows) or (SkinType = sknDelphiStyle) then
        if GetThemedElementDetails(Component, State, Details) then
          GetThemedElementTextColor(Details, Result);
    skncListItem:
      if SkinType <> sknSkin then
        if SkinType = sknDelphiStyle then begin
          {$IF CompilerVersion >= 23} // for Delphi XE2 and up
          if State in [sknsChecked, sknsCheckedAndHotTrack] then
            Result := SpTBXThemeServices.GetStyleFontColor(sfListItemTextSelected)
          else
            Result := SpTBXThemeServices.GetStyleFontColor(sfListItemTextNormal);
          {$IFEND}
        end
        else
          if State in [sknsChecked, sknsCheckedAndHotTrack] then
            Result := clHighlightText;
    skncDockablePanelTitleBar:
      if SkinType = sknSkin then
        Result := GetTextColor(skncToolbarItem, State) // Use skncToolbarItem to get the default text color
      else
        if (SkinType = sknWindows) or (SkinType = sknDelphiStyle) then
          if GetThemedElementDetails(Component, State, Details) then
            GetThemedElementTextColor(Details, Result);
    skncStatusBar:
      if SkinType = sknSkin then
        Result := GetTextColor(skncToolbarItem, State) // Use skncToolbarItem to get the default text color
      else
        if (SkinType = sknWindows) or (SkinType = sknDelphiStyle) then begin
          Details := SpTBXThemeServices.GetElementDetails(tsPane);
          GetThemedElementTextColor(Details, Result);
        end;
    skncTabToolbar:
      if SkinType = sknSkin then
        Result := GetTextColor(skncToolbarItem, State) // Use skncToolbarItem to get the default text color
      else
        if (SkinType = sknWindows) or (SkinType = sknDelphiStyle) then
          Result := GetTextColor(skncTab, State);
    skncWindowTitleBar:
      case SkinType of
        sknNone:
          if State = sknsDisabled then
            Result := clInactiveCaptionText
          else
            Result := clCaptionText;
        sknWindows, sknDelphiStyle:
          if GetThemedElementDetails(Component, State, Details) then
            GetThemedElementTextColor(Details, Result);
        sknSkin:
          Result := GetTextColor(skncToolbarItem, State);  // Use skncToolbarItem to get the default text color
      end;
  end;
end;

function TSpTBXSkinOptions.GetThemedElementDetails(Component: TSpTBXSkinComponentsType;
  Enabled, Pushed, HotTrack, Checked, Focused, Defaulted, Grayed: Boolean; out Details: TThemedElementDetails): Boolean;
var
  SkinType: TSpTBXSkinType;
begin
  Result := False;
  SkinType := SkinManager.GetSkinType;

  case Component of
    skncDock:
      begin
        if SkinType = sknDelphiStyle then
          Details := SpTBXThemeServices.GetElementDetails(ttbToolBarDontCare)
        else
          Details := SpTBXThemeServices.GetElementDetails(trRebarDontCare);
        Result := True;
      end;
    skncDockablePanel:
      begin
        // [Old-Themes]
        {$IF CompilerVersion >= 23}
        // tcpBackground is defined only on XE2 and up
        // Only used with VCL Styles
        Details := SpTBXThemeServices.GetElementDetails(tcpBackground);
        Result := True;
        {$IFEND}
      end;
    skncDockablePanelTitleBar:
      begin
        // [Old-Themes]
        {$IF CompilerVersion >= 23}
        // tcpThemedHeader is defined only on XE2 and up
        // Only used with VCL Styles
        Details := SpTBXThemeServices.GetElementDetails(tcpThemedHeader);
        Result := True;
        {$IFEND}
      end;
    skncGutter:
      begin
        // [Old-Themes]
        {$IF CompilerVersion >= 23} //for Delphi XE2 and up
        Details := SpTBXThemeServices.GetElementDetails(tmPopupGutter);
        {$ELSE}
        Details.Element := teMenu;
        Details.Part := MENU_POPUPGUTTER;
        Details.State := 0;
        {$IFEND}
        Result := True;
      end;
    {
    skncMenuBar: ;
    skncOpenToolbarItem: ;
    }
    skncPanel:
      begin
        if Enabled then
          Details := SpTBXThemeServices.GetElementDetails(tbGroupBoxNormal)
        else
          Details := SpTBXThemeServices.GetElementDetails(tbGroupBoxDisabled);
        Result := True;
      end;
    skncPopup:
      begin
        // [Old-Themes]
        {$IF CompilerVersion >= 23} //for Delphi XE2 and up
        Details := SpTBXThemeServices.GetElementDetails(tmPopupBackground);
        {$ELSE}
        Details.Element := teMenu;
        Details.Part := MENU_POPUPBACKGROUND;
        Details.State := 0;
        {$IFEND}
        Result := True;
      end;
    skncSeparator:
      begin
        if Enabled then  // Enabled = Vertical
          Details := SpTBXThemeServices.GetElementDetails(ttbSeparatorNormal)
        else
          Details := SpTBXThemeServices.GetElementDetails(ttbSeparatorVertNormal);
        Result := True;
      end;
    {
    skncSplitter: ;
    }
    skncStatusBar:
      begin
        Details := SpTBXThemeServices.GetElementDetails(tsStatusRoot);
        Result := True;
      end;
    skncStatusBarGrip:
      begin
        Details := SpTBXThemeServices.GetElementDetails(tsGripper);
        Result := True;
      end;
    {
    skncTabBackground: ;
    skncTabToolbar: ;
    }
    skncToolbar:
      begin
        // [Old-Themes]
        // On older versions is trBandNormal on XE2 is trBand
        {$IF CompilerVersion >= 23} // for Delphi XE2 and up
        Details := SpTBXThemeServices.GetElementDetails(trBand);
        {$ELSE}
        Details := SpTBXThemeServices.GetElementDetails(trBandNormal);
        {$IFEND}
        Result := True;
      end;
    skncToolbarGrip:
      begin
        if Enabled then  // Enabled = Vertical
          Details := SpTBXThemeServices.GetElementDetails(trGripperVert)
        else
          Details := SpTBXThemeServices.GetElementDetails(trGripper);
        Result := True;
      end;
    skncWindow: ;
    skncWindowTitleBar:
      begin
        if SkinType = sknDelphiStyle then begin
          if Enabled then
            Details := SpTBXThemeServices.GetElementDetails(twCaptionActive)
          else
            Details := SpTBXThemeServices.GetElementDetails(twCaptionInActive);
        end
        else
          // On WinXP when twCaptionActive is used instead of twSmallCaptionActive the top borders are rounded
          if Enabled then
            Details := SpTBXThemeServices.GetElementDetails(twSmallCaptionActive)
          else
            Details := SpTBXThemeServices.GetElementDetails(twSmallCaptionInActive);
        Result := True;
      end;
    skncMenuBarItem:
      begin
        if SpIsWinVistaOrUp or (SkinType = sknDelphiStyle) then begin
          // [Old-Themes]
          {$IF CompilerVersion >= 23} //for Delphi XE2 and up
          if not Enabled then Details := SpTBXThemeServices.GetElementDetails(tmMenuBarItemDisabled)
          else if Pushed then Details := SpTBXThemeServices.GetElementDetails(tmMenuBarItemPushed)
          else if HotTrack then Details := SpTBXThemeServices.GetElementDetails(tmMenuBarItemHot)
          else Details := SpTBXThemeServices.GetElementDetails(tmMenuBarItemNormal);
          {$ELSE}
          Details.Element := teMenu;
          Details.Part := MENU_BARITEM;
          if not Enabled then Details.State := MBI_DISABLED
          else if Pushed then Details.State := MBI_PUSHED
          else if HotTrack then Details.State := MBI_HOT
          else Details.State := MBI_NORMAL;
          {$IFEND}
        end
        else
          Details := SpTBXThemeServices.GetElementDetails(tmMenuBarItem);
        Result := True;
      end;
    skncMenuItem:
      begin
        if SpIsWinVistaOrUp or (SkinType = sknDelphiStyle) then begin
          // [Old-Themes]
          {$IF CompilerVersion >= 23} //for Delphi XE2 and up
          if not Enabled and HotTrack then Details := SpTBXThemeServices.GetElementDetails(tmPopupItemDisabledHot)
          else if not Enabled then Details := SpTBXThemeServices.GetElementDetails(tmPopupItemDisabled)
          else if HotTrack then Details := SpTBXThemeServices.GetElementDetails(tmPopupItemHot)
          else Details := SpTBXThemeServices.GetElementDetails(tmPopupItemNormal);
          {$ELSE}
          Details.Element := teMenu;
          Details.Part := MENU_POPUPITEM;
          if not Enabled and HotTrack then Details.State := MPI_DISABLEDHOT
          else if not Enabled then Details.State := MPI_DISABLED
          else if HotTrack then Details.State := MPI_HOT
          else Details.State := MPI_NORMAL;
          {$IFEND}
        end
        else begin
          if HotTrack then
            Details := SpTBXThemeServices.GetElementDetails(tmMenuItemSelected)
          else
            Details := SpTBXThemeServices.GetElementDetails(tmMenuItemNormal);
        end;
        Result := True;
      end;
    skncToolbarItem:
      begin
        if not Enabled then Details := SpTBXThemeServices.GetElementDetails(ttbButtonDisabled)
        else if Pushed then Details := SpTBXThemeServices.GetElementDetails(ttbButtonPressed)
        else if HotTrack and Checked then Details := SpTBXThemeServices.GetElementDetails(ttbButtonCheckedHot)
        else if HotTrack then Details := SpTBXThemeServices.GetElementDetails(ttbButtonHot)
        else if Checked then Details := SpTBXThemeServices.GetElementDetails(ttbButtonChecked)
        else Details := SpTBXThemeServices.GetElementDetails(ttbButtonNormal);
        Result := True;
      end;
    skncButton, skncEditButton:
      begin
        if not Enabled then Details := SpTBXThemeServices.GetElementDetails(tbPushButtonDisabled)
        else if Pushed or Checked then Details := SpTBXThemeServices.GetElementDetails(tbPushButtonPressed)
        else if HotTrack then Details := SpTBXThemeServices.GetElementDetails(tbPushButtonHot)
        else if Defaulted or Focused then Details := SpTBXThemeServices.GetElementDetails(tbPushButtonDefaulted)
        else Details := SpTBXThemeServices.GetElementDetails(tbPushButtonNormal);
        Result := True;
      end;
    skncCheckBox:
      begin
        if Grayed then begin
          if not Enabled then Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxMixedDisabled)
          else if Pushed then Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxMixedPressed)
          else if HotTrack then Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxMixedHot)
          else Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxMixedNormal);
        end
        else
          if Checked then begin
            if not Enabled then Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxCheckedDisabled)
            else if Pushed then Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxCheckedPressed)
            else if HotTrack then Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxCheckedHot)
            else Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxCheckedNormal);
          end
          else begin
            if not Enabled then Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxUncheckedDisabled)
            else if Pushed then Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxUncheckedPressed)
            else if HotTrack then Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxUncheckedHot)
            else Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxUncheckedNormal);
          end;
        Result := True;
      end;
    skncEditFrame:
      begin
        // [Old-Themes]
        {$IF CompilerVersion >= 23} //for Delphi XE2 and up
        if not SpIsWinVistaOrUp then Details := SpTBXThemeServices.GetElementDetails(tcComboBoxDontCare)
        else if not Enabled then Details := SpTBXThemeServices.GetElementDetails(tcBorderDisabled)
        else if HotTrack then Details := SpTBXThemeServices.GetElementDetails(tcBorderHot)
        else Details := SpTBXThemeServices.GetElementDetails(tcBorderNormal);
        {$ELSE}
        Details.Element := teComboBox;
        if SpIsWinVistaOrUp then begin
          // Use the new API on Windows Vista
          Details.Part := CP_BORDER;
          if not Enabled then Details.State := CBXS_DISABLED
          else if HotTrack then Details.State := CBXS_HOT
          else Details.State := CBXS_NORMAL;
        end
        else begin
          Details.Part := 0;
          Details.State := 0;
        end;
        {$IFEND}
        Result := True;
      end;
    skncHeader:
      begin
        // [Old-Themes]
        {$IF CompilerVersion >= 23} //for Delphi XE2 and up
        if Pushed then Details := SpTBXThemeServices.GetElementDetails(thHeaderItemPressed)
        else if HotTrack then Details := SpTBXThemeServices.GetElementDetails(thHeaderItemHot)
        else Details := SpTBXThemeServices.GetElementDetails(thHeaderItemNormal);
        {$ELSE}
        Details.Element := teHeader;
        Details.Part := HP_HEADERITEM;
        if Pushed then Details.State := HIS_PRESSED
        else if HotTrack then Details.State := HIS_HOT
        else Details.State := HIS_NORMAL;
        {$IFEND}
        Result := True;
      end;
    skncLabel:
      begin
        if Enabled then
          Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxUncheckedNormal)
        else
          Details := SpTBXThemeServices.GetElementDetails(tbCheckBoxUncheckedDisabled);
        Result := True;
      end;
    {
    skncListItem: ;
    }
    skncRadioButton:
      begin
        if Checked then begin
          if not Enabled then Details := SpTBXThemeServices.GetElementDetails(tbRadioButtonCheckedDisabled)
          else if Pushed then Details := SpTBXThemeServices.GetElementDetails(tbRadioButtonCheckedPressed)
          else if HotTrack then Details := SpTBXThemeServices.GetElementDetails(tbRadioButtonCheckedHot)
          else Details := SpTBXThemeServices.GetElementDetails(tbRadioButtonCheckedNormal);
        end
        else begin
          if not Enabled then Details := SpTBXThemeServices.GetElementDetails(tbRadioButtonUncheckedDisabled)
          else if Pushed then Details := SpTBXThemeServices.GetElementDetails(tbRadioButtonUncheckedPressed)
          else if HotTrack then Details := SpTBXThemeServices.GetElementDetails(tbRadioButtonUncheckedHot)
          else Details := SpTBXThemeServices.GetElementDetails(tbRadioButtonUncheckedNormal);
        end;
        Result := True;
      end;
    skncTab:
      begin
        if not Enabled then Details := SpTBXThemeServices.GetElementDetails(ttTabItemDisabled)
        else if Pushed or (HotTrack and Checked) or Checked then Details := SpTBXThemeServices.GetElementDetails(ttTabItemSelected)
        else if HotTrack then Details := SpTBXThemeServices.GetElementDetails(ttTabItemHot)
        else Details := SpTBXThemeServices.GetElementDetails(ttTabItemNormal);
        Result := True;
      end;
    skncProgressBar:    ; // not used
    skncTrackBar:       ; // not used
    skncTrackBarButton: ; // not used
  end;
end;

function TSpTBXSkinOptions.GetThemedElementDetails(Component: TSpTBXSkinComponentsType;
  State: TSpTBXSkinStatesType; out Details: TThemedElementDetails): Boolean;
var
  Enabled, Pushed, HotTrack, Checked: Boolean;
begin
  GetState(State, Enabled, Pushed, HotTrack, Checked);
  Result := GetThemedElementDetails(Component, Enabled, Pushed, HotTrack, Checked, False, False, False, Details);
end;

function TSpTBXSkinOptions.GetThemedElementSize(ACanvas: TCanvas; Details: TThemedElementDetails; DPI: Integer): TSize;
begin
  {$IF CompilerVersion >= 23} // for Delphi XE2 and up
  SpTBXThemeServices.GetElementSize(ACanvas.Handle, Details, esActual, Result{$IF CompilerVersion >= 33}, DPI{$IFEND}); // DPI param introduced on 10.3 Rio
  {$ELSE}
  GetThemePartSize(SpTBXThemeServices.Theme[Details.Element], ACanvas.Handle, Details.Part, Details.State, nil, TS_TRUE, Result);
  {$IFEND}
end;

procedure TSpTBXSkinOptions.GetThemedElementTextColor(Details: TThemedElementDetails; out AColor: TColor);
begin
  {$IF CompilerVersion >= 23} // for Delphi XE2 and up
  SpTBXThemeServices.GetElementColor(Details, ecTextColor, AColor);
  {$ELSE}
  GetThemeColor(SpTBXThemeServices.Theme[Details.Element], Details.Part, Details.State,
      TMT_TEXTCOLOR, TColorRef(AColor));
  {$IFEND}
end;

function TSpTBXSkinOptions.GetThemedSystemColor(AColor: TColor): TColor;
begin
  {$IF CompilerVersion >= 23} // for Delphi XE2 and up
  Result := SpTBXThemeServices.GetSystemColor(AColor);
  {$ELSE}
  Result := AColor;
  {$IFEND}
end;

procedure TSpTBXSkinOptions.PaintBackground(ACanvas: TCanvas; ARect: TRect;
  Component: TSpTBXSkinComponentsType; State: TSpTBXSkinStatesType;
  Background, Borders: Boolean; Vertical: Boolean = False;
  ForceRectBorders: TAnchors = []);
var
  BackgroundRect: TRect;
  Op: TSpTBXSkinOptionCategory;
begin
  Op := Options(Component, State);

  if Op.Borders.IsEmpty then
    Borders := False;
  if Op.Body.IsEmpty then
    Background := False;

  if Background then begin
    BackgroundRect := ARect;
    if Borders then
      InflateRect(BackgroundRect, -1, -1);
    SpPaintSkinBackground(ACanvas, BackgroundRect, Op, Vertical);
  end;

  if Borders then
    SpPaintSkinBorders(ACanvas, ARect, Op, ForceRectBorders);
end;

procedure TSpTBXSkinOptions.PaintThemedElementBackground(ACanvas: TCanvas;
  ARect: TRect; Details: TThemedElementDetails; DPI: Integer);
var
  SaveIndex: Integer;
begin
  SaveIndex := SaveDC(ACanvas.Handle);  // XE2 Styles changes the font
  try
    SpTBXThemeServices.DrawElement(ACanvas.Handle, Details, ARect, nil{$IF CompilerVersion >= 33}, DPI{$IFEND}); // DPI param introduced on 10.3 Rio DPI);
  finally
    RestoreDC(ACanvas.Handle, SaveIndex);
  end;
end;

procedure TSpTBXSkinOptions.PaintThemedElementBackground(ACanvas: TCanvas;
  ARect: TRect; Component: TSpTBXSkinComponentsType; Enabled, Pushed, HotTrack,
  Checked, Focused, Defaulted, Grayed: Boolean; DPI: Integer);
var
  Details: TThemedElementDetails;
begin
  if GetThemedElementDetails(Component, Enabled, Pushed, HotTrack, Checked, Focused, Defaulted, Grayed, Details) then
    PaintThemedElementBackground(ACanvas, ARect, Details, DPI);
end;

procedure TSpTBXSkinOptions.PaintThemedElementBackground(ACanvas: TCanvas;
  ARect: TRect; Component: TSpTBXSkinComponentsType;
  State: TSpTBXSkinStatesType; DPI: Integer);
var
  Details: TThemedElementDetails;
begin
  if GetThemedElementDetails(Component, State, Details) then
    PaintThemedElementBackground(ACanvas, ARect, Details, DPI);
end;

procedure TSpTBXSkinOptions.PaintMenuCheckMark(ACanvas: TCanvas; ARect: TRect;
  Checked, Grayed: Boolean; State: TSpTBXSkinStatesType; DPI: Integer);
var
  CheckColor: TColor;
  VistaCheckSize: TSize;
  Details: TThemedElementDetails;
  SkinType: TSpTBXSkinType;
begin
  SkinType := SkinManager.GetSkinType;
  // VCL Styles does not DPI scale menu checkmarks, Windows does
  if ((SkinType = sknWindows) and SpIsWinVistaOrUp) or (SkinType = sknDelphiStyle) then
  begin
    // [Old-Themes]
    {$IF CompilerVersion >= 23} //for Delphi XE2 and up
    if State = sknsDisabled then Details := SpTBXThemeServices.GetElementDetails(tmPopupCheckDisabled)
    else Details := SpTBXThemeServices.GetElementDetails(tmPopupCheckNormal);
    {$ELSE}
    Details.Element := teMenu;
    Details.Part := MENU_POPUPCHECK;
    if State = sknsDisabled then Details.State := MC_CHECKMARKDISABLED
    else Details.State := MC_CHECKMARKNORMAL;
    {$IFEND}
    VistaCheckSize := GetThemedElementSize(ACanvas, Details, DPI); // Returns a scaled value
    ARect := SpCenterRect(ARect, VistaCheckSize.cx, VistaCheckSize.cy);
    PaintThemedElementBackground(ACanvas, ARect, Details, DPI);
  end
  else begin
    if SkinType = sknNone then
      CheckColor := clMenuText // On sknNone it's clMenuText even when disabled
    else
      CheckColor := GetTextColor(skncMenuItem, State);
    SpDrawGlyphPattern(ACanvas, ARect, gptMenuCheckmark, CheckColor, DPI);
  end;
end;

procedure TSpTBXSkinOptions.PaintMenuRadioMark(ACanvas: TCanvas; ARect: TRect;
  Checked: Boolean; State: TSpTBXSkinStatesType; DPI: Integer);
var
  CheckColor: TColor;
  VistaCheckSize: TSize;
  Details: TThemedElementDetails;
  SkinType: TSpTBXSkinType;
begin
  SkinType := SkinManager.GetSkinType;
  // Vcl Styles does not DPI scale menu radio buttons, Windows does
  if ((SkinType = sknWindows) and SpIsWinVistaOrUp) or
     ((SkinType = sknDelphiStyle) and (Screen.PixelsPerInch = 96)) then
  begin
    // [Old-Themes]
    {$IF CompilerVersion >= 23} //for Delphi XE2 and up
    if State = sknsDisabled then Details := SpTBXThemeServices.GetElementDetails(tmPopupBulletDisabled)
    else Details := SpTBXThemeServices.GetElementDetails(tmPopupBulletNormal);
    {$ELSE}
    Details.Element := teMenu;
    Details.Part := MENU_POPUPCHECK;
    if State = sknsDisabled then Details.State := MC_BULLETDISABLED
    else Details.State := MC_BULLETNORMAL;
    {$IFEND}
    VistaCheckSize := GetThemedElementSize(ACanvas, Details, DPI); // Returns a scaled value
    ARect := SpCenterRect(ARect, VistaCheckSize.cx, VistaCheckSize.cy);
    PaintThemedElementBackground(ACanvas, ARect, Details, DPI);
  end
  else begin
    if SkinType = sknNone then
      CheckColor := clMenuText // On sknNone it's clMenuText even when disabled
    else
      CheckColor := GetTextColor(skncMenuItem, State);
    SpDrawGlyphPattern(ACanvas, ARect, gptMenuRadiomark, CheckColor, DPI);
  end;
end;

procedure TSpTBXSkinOptions.PaintWindowFrame(ACanvas: TCanvas; ARect: TRect;
  IsActive, DrawBody: Boolean; BorderSize: Integer = 4);
var
  C: TColor;
  R: TRect;
  I: Integer;
  State: TSpTBXSkinStatesType;
  Op: TSpTBXSkinOptionEntry;
begin
  if IsActive then
    State := sknsNormal
  else
    if Options(skncWindow, sknsDisabled).IsEmpty then
      State := sknsNormal
    else
      State := sknsDisabled;

  C := ACanvas.Brush.Color;
  if DrawBody then
    PaintBackground(ACanvas, ARect, skncWindow, State, True, False);
  R := ARect;
  Op := Options(skncWindow, State).Borders;
  for I := 1 to BorderSize do begin
    if I = 1 then ACanvas.Brush.Color := Op.Color1
    else if I = 2 then ACanvas.Brush.Color := Op.Color2
    else if I = 3 then ACanvas.Brush.Color := Op.Color3
    else if I >= 4 then ACanvas.Brush.Color := Op.Color4;
    ACanvas.FrameRect(R);
    InflateRect(R, -1, -1);
  end;
  ACanvas.Brush.Color := C;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXSkinsListEntry }

destructor TSpTBXSkinsListEntry.Destroy;
begin
  SkinClass := nil;
  FreeAndNil(SkinStrings);

  inherited;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXSkinManager }

constructor TSpTBXSkinManager.Create;
begin
  FNotifies := TList.Create;
  FCurrentSkin := TSpTBXSkinOptions.Create;
  FSkinsList := TSpTBXSkinsDictionary.Create([doOwnsValues]);
end;

destructor TSpTBXSkinManager.Destroy;
begin
  FreeAndNil(FNotifies);
  FreeAndNil(FCurrentSkin);
  FreeAndNil(FSkinsList);
  inherited;
end;

procedure TSpTBXSkinManager.AddSkin(SkinName: string;
  SkinClass: TSpTBXSkinOptionsClass);
var
  K: TSpTBXSkinsListEntry;
begin
  if not FSkinsList.ContainsKey(SkinName) then begin
    K := TSpTBXSkinsListEntry.Create;
    try
      K.SkinClass := SkinClass;
      FSkinsList.Add(SkinName, K);  // the list owns K
    except
      K.Free;
    end;
  end;
end;

procedure TSpTBXSkinManager.AddSkin(SkinOptions: TStrings);
var
  K: TSpTBXSkinsListEntry;
  S: string;
begin
  K := TSpTBXSkinsListEntry.Create;
  try
    K.SkinStrings := TStringList.Create;
    S := SkinOptions.Values['Name '];
    if S = '' then
      S := SkinOptions.Values['Name'];
    S := Trim(S);
    if (S <> '') and not FSkinsList.ContainsKey(S) then begin
      K.SkinStrings.Assign(SkinOptions);
      FSkinsList.Add(S, K);  // the list owns K
    end
    else
      K.Free;
  except
    K.Free;
  end;
end;

function TSpTBXSkinManager.AddSkinFromFile(Filename: string): string;
var
  L: TStringList;
begin
  L := TStringList.Create;
  try
    L.LoadFromFile(Filename);
    AddSkin(L);

    Result := L.Values['Name '];
    if Result = '' then
      Result := L.Values['Name'];
    Result := Trim(Result);
  finally
    L.Free;
  end;
end;

procedure TSpTBXSkinManager.AddSkinNotification(AObject: TObject);
begin
  if FNotifies.IndexOf(AObject) < 0 then FNotifies.Add(AObject);
end;

procedure TSpTBXSkinManager.RemoveSkinNotification(AObject: TObject);
begin
  FNotifies.Remove(AObject);
end;

procedure TSpTBXSkinManager.ResetToSystemStyle;
begin
  {$IF CompilerVersion >= 23} // for Delphi XE2 and up
  // Reset the VCL Style to System Style
  if TStyleManager.IsCustomStyleActive then
    TStyleManager.SetStyle(TStyleManager.SystemStyle);
  {$IFEND}
end;

procedure TSpTBXSkinManager.Broadcast;
var
  Msg: TMessage;
  I: Integer;
begin
  if FNotifies.Count > 0 then begin
    Msg.Msg := WM_SPSKINCHANGE;
    Msg.WParam := 0;
    Msg.LParam := 0;
    Msg.Result := 0;
    for I := 0 to FNotifies.Count - 1 do
      TObject(FNotifies[I]).Dispatch(Msg);
  end;
  if Assigned(FOnSkinChange) then FOnSkinChange(Self);
end;

procedure TSpTBXSkinManager.BroadcastSkinNotification;
begin
  Broadcast;
end;

procedure TSpTBXSkinManager.LoadFromFile(Filename: string);
begin
  FCurrentSkin.LoadFromFile(Filename);
end;

procedure TSpTBXSkinManager.SaveToFile(Filename: string);
begin
  FCurrentSkin.SaveToFile(Filename);
end;

function TSpTBXSkinManager.GetCurrentSkinName: string;
begin
  Result := FCurrentSkin.SkinName;
end;

function TSpTBXSkinManager.GetSkinType: TSpTBXSkinType;
begin
  Result := sknSkin;

  {$IF CompilerVersion >= 23} // for Delphi XE2 and up
  if TStyleManager.IsCustomStyleActive then
    Result := sknDelphiStyle;
  {$IFEND}

  if (Result = sknSkin) and IsDefaultSkin then
    Result := sknWindows;

  if (Result = sknWindows) and not SkinManager.IsXPThemesEnabled then
    Result := sknNone;
end;

procedure TSpTBXSkinManager.GetSkinsAndDelphiStyles(SkinsAndStyles: TStrings);
var
  I: Integer;
  S: string;
  L: TStringList;
begin
  L := TStringList.Create;
  try
    SkinsAndStyles.Clear;

    // Fill Skins
    L.Add('Default');
    for S in SkinManager.SkinsList.Keys do
      L.Add(S);

    // Fill Styles
    {$IF CompilerVersion >= 23} // for Delphi XE2 and up
    for S in TStyleManager.StyleNames do
      if S <> TStyleManager.SystemStyle.Name then
        L.Add(S);
    {$IFEND}

    // Sort the list and move the Default skin to the top
    L.Sort;
    I := L.IndexOf('Default');
    if I > -1 then L.Move(I, 0);
    SkinsAndStyles.AddStrings(L);
  finally
    L.Free;
  end;
end;

function TSpTBXSkinManager.IsDefaultSkin: Boolean;
begin
  // Returns True if CurrentSkin is the default Windows theme
  // For Delphi XE2 and up make sure setting System style sets CurrentSkin to Default
  Result := CurrentSkinName = 'Default';
end;

function TSpTBXSkinManager.IsXPThemesEnabled: Boolean;
begin
  {$IF CompilerVersion >= 23} //for Delphi XE2 and up
  Result := StyleServices.Enabled;
  {$ELSE}
  Result := ThemeServices.ThemesAvailable and UxTheme.UseThemes;
  {$IFEND}
end;

procedure TSpTBXSkinManager.SetSkin(SkinName: string);
var
  K: TSpTBXSkinsListEntry;
begin
  if not SameText(SkinName, CurrentSkinName) or (GetSkinType = sknDelphiStyle) then
    if (SkinName = '') or SameText(SkinName, 'Default') then
      SetToDefaultSkin
    else begin
      if FSkinsList.TryGetValue(SkinName, K) then
        if Assigned(K.SkinClass) then begin
          FCurrentSkin.Free;
          FCurrentSkin := K.SkinClass.Create;
          ResetToSystemStyle;
          BroadcastSkinNotification;
        end
        else
          if Assigned(K.SkinStrings) then begin
            FCurrentSkin.Free;
            FCurrentSkin := TSpTBXSkinOptions.Create;
            FCurrentSkin.LoadFromStrings(K.SkinStrings);
            ResetToSystemStyle;
            BroadcastSkinNotification;
          end;
    end;
end;

// [Old-Themes]
{$IF CompilerVersion >= 23} // for Delphi XE2 and up
procedure TSpTBXSkinManager.SetDelphiStyle(StyleName: string);
begin
  if not SameText(StyleName, TStyleManager.ActiveStyle.Name) then begin
    if not IsDefaultSkin then
      SetToDefaultSkin;
    TStyleManager.SetStyle(StyleName);
    Broadcast;
  end;
end;

function TSpTBXSkinManager.IsValidDelphiStyle(StyleName: string): Boolean;
var
  S: string;
begin
  Result := False;
  for S in TStyleManager.StyleNames do
    if SameText(S, StyleName) then begin
      Result := True;
      Exit;
    end;
end;
{$IFEND}

procedure TSpTBXSkinManager.SetToDefaultSkin;
begin
  if GetSkinType <> sknWindows then begin
//  if not IsDefaultSkin or IsDelphiStyleActive then begin
    FCurrentSkin.Free;
    FCurrentSkin := TSpTBXSkinOptions.Create;
    ResetToSystemStyle;
    BroadcastSkinNotification;
  end;
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXSkinSwitcher }

constructor TSpTBXSkinSwitcher.Create(AOwner: TComponent);
begin
  inherited;
  SkinManager.AddSkinNotification(Self);
end;

destructor TSpTBXSkinSwitcher.Destroy;
begin
  SkinManager.RemoveSkinNotification(Self);
  inherited;
end;

function TSpTBXSkinSwitcher.GetSkin: string;
begin
  Result := SkinManager.CurrentSkinName;
end;

procedure TSpTBXSkinSwitcher.SetSkin(const Value: string);
begin
  SkinManager.SetSkin(Value);
end;

procedure TSpTBXSkinSwitcher.WMSpSkinChange(var Message: TMessage);
begin
  if Assigned(FOnSkinChange) then FOnSkinChange(Self);
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Stock Objects }

procedure InitializeStock;
begin
  StockBitmap := TBitmap.Create;
  StockBitmap.SetSize(8, 8);

  @SpPrintWindow := GetProcAddress(GetModuleHandle(user32), 'PrintWindow');

  if not Assigned(FInternalSkinManager) then
    FInternalSkinManager := TSpTBXSkinManager.Create;
end;

procedure FinalizeStock;
begin
  FreeAndNil(StockBitmap);
  FreeAndNil(FInternalSkinManager);
end;

initialization
  InitializeStock;

finalization
  FinalizeStock;

end.

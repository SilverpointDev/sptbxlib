unit SpTBXImageList;

{==============================================================================
Version 2.5.8

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

Development notes:
  - All the Windows and Delphi bugs fixes are marked with '[Bugfix]'.
  - All the theme changes and adjustments are marked with '[Theme-Change]'.
  - Not included in the lib, only used by the demos for now.
  - To make unit part of the lib add vclimg (uses PngImage) and vclwinx (uses
    ImageCollection and VirtualImageList units) to the required packages, for
    Delphi Rio and up. Need to make 2 separate packages: one for older versions
    of Delphi and the other for Rio and up.

==============================================================================}

interface

{$BOOLEVAL OFF}   // Unit depends on short-circuit boolean evaluation
{$IF CompilerVersion >= 25} // for Delphi XE4 and up
  {$LEGACYIFEND ON} // requires $IF to be terminated with $ENDIF instead of $IFEND
{$IFEND}

uses
  Windows, Messages, Classes, SysUtils, Graphics,
  ImgList,
  {$IF CompilerVersion >= 24} // for Delphi XE3 and up
  System.UITypes,
  {$IFEND}
  {$IF CompilerVersion >= 33} // for Delphi Rio and up
  // TImageCollection and TVirtualImagelist introduced on Rio
  Vcl.VirtualImageList, Vcl.BaseImageCollection, Vcl.ImageCollection,
  {$IFEND}
  Types;

type
  { TSpTBXImageList }

  TSpTBXImageList = class
  public
    ImageList: TCustomImageList;
    {$IF CompilerVersion >= 33} // for Delphi Rio and up
    ImageCollection: TImageCollection;
    {$IFEND}
    constructor Create(AOwner: TComponent); virtual;
    destructor Destroy; override;
    procedure LoadGlyphs(GlyphPath: string);
  end;

{ Utils }
procedure SpLoadGlyphs(IL: TCustomImageList; GlyphPath: string);

implementation

uses
  TB2Common, IOUtils, pngimage, Generics.Collections, Generics.Defaults;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ Utils }

procedure SpLoadGlyphs(IL: TCustomImageList; GlyphPath: string);
// Finds png files on GlyphPath and adds them to IL
// If IL is a TVirtualImageList it adds all the PNGs sizes.
// Otherwise it adds only the PNGs that matches the size of the IL
// Notation of files must be filename-16x16.png
var
  Files: TStringDynArray;
  FilenameS, S, ILSize: string;
  I: Integer;
  P: TPngImage;
  B: TBitmap;
begin
  ILSize := Format('%dX%d', [IL.Width, IL.Height]);
  Files := TDirectory.GetFiles(GlyphPath, '*.png');
  TArray.Sort<string>(Files, TStringComparer.Ordinal);

  for S in Files do begin
    {$IF CompilerVersion >= 33} // for Delphi Rio and up
    // TImageCollection and TVirtualImagelist introduced on Rio
    if IL is TVirtualImageList then begin
      if Assigned(TVirtualImageList(IL).ImageCollection) and (TVirtualImageList(IL).ImageCollection is TImageCollection) then begin
        FilenameS := TPath.GetFileName(S);
        I := LastDelimiter('-_', FilenameS);
        if I > 1 then
          FilenameS := Copy(FilenameS, 1, I-1);
        // Add all the sizes of the png with 1 name on ImageCollection
        TImageCollection(TVirtualImageList(IL).ImageCollection).Add(FilenameS, S);
      end;
    end
    else
    {$IFEND}
    begin
      // Try to add only PNGs with the same size as the Image List
      // Notation of files must be filename-16x16.png
      FilenameS := TPath.GetFileNameWithoutExtension(S);
      I := LastDelimiter('-_', FilenameS) + 1;
      if I > 2 then begin
        FilenameS := Copy(FilenameS, I, Length(ILSize));
        if UpperCase(FilenameS) <> ILSize then
          FilenameS := '';
      end;
      if FilenameS <> '' then begin
        P := TPNGImage.Create;
        B := TBitmap.Create;
        try
          P.LoadFromFile(S);
          B.Assign(P);
          IL.ColorDepth := cd32Bit;
          IL.Add(B, nil);
        finally
          P.Free;
          B.Free;
        end;
      end;
    end;
  end;

  {$IF CompilerVersion >= 33} // for Delphi Rio and up
  if IL is TVirtualImageList then
    TVirtualImageList(IL).AutoFill := True;
  {$IFEND}
end;

//WMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWMWM
{ TSpTBXImageList }

constructor TSpTBXImageList.Create(AOwner: TComponent);
begin
  {$IF CompilerVersion >= 33} // for Delphi Rio and up
  // TImageCollection and TVirtualImagelist introduced on Rio
  ImageCollection := TImageCollection.Create(AOwner);
  ImageList := TVirtualImageList.Create(AOwner);
  TVirtualImageList(ImageList).ImageCollection := ImageCollection;
  {$ELSE}
  ImageList := TImageList.Create(Self);
  {$IFEND}
end;

destructor TSpTBXImageList.Destroy;
begin
  {$IF CompilerVersion >= 33} // for Delphi Rio and up
  // TImageCollection and TVirtualImagelist introduced on Rio
  ImageCollection.Free;
  {$IFEND}
  ImageList.Free;

  inherited;
end;

procedure TSpTBXImageList.LoadGlyphs(GlyphPath: string);
begin
  SpLoadGlyphs(ImageList, GlyphPath);
end;

end.

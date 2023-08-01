object SpTBXCustomizeForm: TSpTBXCustomizeForm
  Left = 219
  Top = 115
  Caption = 'Customize...'
  ClientHeight = 320
  ClientWidth = 343
  Color = clBtnFace
  Constraints.MinHeight = 355
  Constraints.MinWidth = 355
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Shell Dlg 2'
  Font.Style = []
  FormStyle = fsStayOnTop
  KeyPreview = True
  Position = poDesigned
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  OnShow = FormShow
  TextHeight = 13
  object SpTBXTabControl1: TSpTBXTabControl
    Left = 0
    Top = 0
    Width = 343
    Height = 285
    Align = alClient
    ActiveTabIndex = 0
    HiddenItems = <>
    object tabToolbars: TSpTBXTabItem
      Caption = 'Toolbars'
      Checked = True
    end
    object tabCommands: TSpTBXTabItem
      Caption = 'Commands'
    end
    object tabShortcuts: TSpTBXTabItem
      Caption = 'Shortcuts'
    end
    object SpTBXTabSheet2: TSpTBXTabSheet
      Left = 0
      Top = 25
      Width = 347
      Height = 261
      Caption = 'Shortcuts'
      ImageIndex = -1
      ExplicitWidth = 343
      ExplicitHeight = 260
      DesignSize = (
        343
        260)
      TabItem = 'tabShortcuts'
      object SpTBXPanel5: TSpTBXPanel
        Left = 8
        Top = 226
        Width = 231
        Height = 22
        Anchors = [akLeft, akRight, akBottom]
        TabOrder = 1
        HotTrack = True
        ExplicitWidth = 235
        object Panel1: TPanel
          Left = 2
          Top = 2
          Width = 231
          Height = 18
          Align = alClient
          BevelOuter = bvNone
          TabOrder = 0
          DesignSize = (
            227
            18)
          object HotKey1: THotKey
            Left = -2
            Top = -2
            Width = 233
            Height = 21
            Anchors = [akLeft, akTop, akRight, akBottom]
            HotKey = 0
            Modifiers = []
            TabOrder = 0
          end
        end
      end
      object ChangeShortcut: TSpTBXButton
        Left = 245
        Top = 225
        Width = 75
        Height = 25
        Caption = 'C&hange'
        Anchors = [akRight, akBottom]
        TabOrder = 2
        OnClick = ChangeShortcutClick
        ExplicitLeft = 249
      end
      object lbShortcuts: TSpTBXListBox
        Left = 8
        Top = 8
        Width = 314
        Height = 207
        Anchors = [akLeft, akTop, akRight, akBottom]
        TabOrder = 0
        OnClick = lbShortcutsClick
        OnDrawItem = lbShortcutsDrawItem
      end
    end
    object SpTBXTabSheet3: TSpTBXTabSheet
      Left = 0
      Top = 25
      Width = 343
      Height = 260
      Caption = 'Commands'
      ImageIndex = -1
      DesignSize = (
        343
        260)
      TabItem = 'tabCommands'
      object SpTBXLabel3: TSpTBXLabel
        Left = 8
        Top = 212
        Width = 314
        Height = 44
        Caption = 
          'To add command buttons, drag and drop commands onto a toolbar. T' +
          'o remove command buttons, drag them off the toolbar and drop the' +
          'm on the commands list.'
        Anchors = [akLeft, akRight, akBottom]
        AutoSize = False
        Wrapping = twWrap
        ExplicitWidth = 318
      end
      object lbCommands: TSpTBXListBox
        Left = 8
        Top = 8
        Width = 314
        Height = 201
        Anchors = [akLeft, akTop, akRight, akBottom]
        DragMode = dmAutomatic
        TabOrder = 0
        OnDragDrop = lbCommandsDragDrop
        OnDragOver = lbCommandsDragOver
        OnDrawItem = lbCommandsDrawItem
        OnEndDrag = lbCommandsEndDrag
        OnStartDrag = lbCommandsStartDrag
      end
    end
    object SpTBXTabSheet1: TSpTBXTabSheet
      Left = 0
      Top = 25
      Width = 343
      Height = 260
      Caption = 'Toolbars'
      ImageIndex = -1
      DesignSize = (
        343
        260)
      TabItem = 'tabToolbars'
      object SpTBXGroupBox1: TSpTBXGroupBox
        Left = 159
        Top = 2
        Width = 161
        Height = 190
        Caption = 'Options'
        Anchors = [akTop, akRight, akBottom]
        TabOrder = 1
        ExplicitLeft = 163
        object cbText: TSpTBXComboBox
          Left = 6
          Top = 66
          Width = 149
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 1
          OnClick = cbTextClick
          Items.Strings = (
            'Icons + Selective Text'
            'Icons'
            'Icons + Text'
            'Text')
        end
        object cbIcon: TSpTBXComboBox
          Left = 6
          Top = 114
          Width = 149
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 2
          Items.Strings = (
            'Large images'
            'Small images')
        end
        object cbTextLabel: TSpTBXLabel
          Left = 6
          Top = 45
          Width = 68
          Height = 19
          Caption = 'T&ext Options'
          FocusControl = cbText
        end
        object cbIconLabel: TSpTBXLabel
          Left = 6
          Top = 95
          Width = 67
          Height = 19
          Caption = '&Icon Options'
          FocusControl = cbIcon
        end
        object checkVisible: TSpTBXCheckBox
          Left = 6
          Top = 21
          Width = 57
          Height = 25
          Caption = '&Visible'
          TabOrder = 0
          OnClick = checkVisibleClick
        end
      end
      object SpTBXGroupBox2: TSpTBXGroupBox
        Left = 159
        Top = 199
        Width = 161
        Height = 49
        Caption = '&Skins'
        Anchors = [akRight, akBottom]
        TabOrder = 2
        ExplicitLeft = 163
        object cbSkins: TSpTBXComboBox
          Left = 6
          Top = 20
          Width = 149
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 0
          OnClick = cbSkinsClick
        end
      end
      object lbToolbars: TSpTBXCheckListBox
        Left = 8
        Top = 8
        Width = 142
        Height = 240
        Anchors = [akLeft, akTop, akRight, akBottom]
        TabOrder = 0
        OnClick = lbToolbarsClick
        OnClickCheck = lbToolbarsClickCheck
      end
    end
  end
  object ClosePanel: TSpTBXPanel
    Left = 0
    Top = 285
    Width = 343
    Height = 35
    Align = alBottom
    TabOrder = 1
    Borders = False
    TBXStyleBackground = True
    DesignSize = (
      343
      35)
    object CloseButton: TSpTBXButton
      Left = 248
      Top = 4
      Width = 78
      Height = 25
      Caption = '&Close'
      Anchors = [akRight, akBottom]
      TabOrder = 0
      OnClick = CloseButtonClick
      ExplicitLeft = 252
    end
    object ResetButton: TSpTBXButton
      Left = 7
      Top = 4
      Width = 130
      Height = 25
      Caption = '&Default Options'
      TabOrder = 1
      OnClick = ResetButtonClick
    end
  end
end
